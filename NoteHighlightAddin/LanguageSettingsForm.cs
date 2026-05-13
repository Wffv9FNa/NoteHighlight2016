/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Helper;

namespace NoteHighlightAddin
{
    /// <summary>
    /// Modal language picker. Hosts a <see cref="DataGridView"/> with two checkbox columns
    /// (Pinned / Enabled) plus the language label and its underlying highlight tag. Up / Down
    /// buttons reorder the pinned rows. OK / Apply / Cancel / Reset to defaults follow the
    /// behaviour in plan section 5.
    ///
    /// <para>
    /// Editing is purely in-memory until Apply / OK. Save is funnelled through
    /// <see cref="SettingsHelper.SaveLanguages"/> -> <see cref="LanguageSettings.SaveAtomically"/>
    /// (named-mutex + atomic write). On successful save, the form calls
    /// <see cref="AddIn.ReplaceLanguageSettings"/> which sets the deferred-invalidate flag; the
    /// next ribbon onAction observes it and calls IRibbonUI.Invalidate(). The form NEVER touches
    /// <c>_ribbon</c> directly - see plan section 6.
    /// </para>
    /// </summary>
    public partial class LanguageSettingsForm : Form
    {
        public const int MaxPinned = 14;

        private readonly AddIn _addin;
        private readonly string _jsonPath;
        private readonly IReadOnlyList<LanguageDescriptor> _registry;

        // Internal model. _pinnedOrder is the ordered pinned list; _enabled holds every enabled
        // tag (pinned implies enabled, so pinned tags are members). _allTags is every tag known
        // to the registry, so the grid shows everything declared in ribbon.xml regardless of
        // current pinned / enabled state. We render pinned rows first (in order) then the rest.
        private readonly List<string> _pinnedOrder = new List<string>();
        private readonly HashSet<string> _enabled = new HashSet<string>(StringComparer.Ordinal);
        private readonly List<LanguageDescriptor> _allTags = new List<LanguageDescriptor>();

        // Suppresses re-entry from CellValueChanged while we mutate the grid programmatically
        // (e.g. ticking E automatically when P is ticked, or reordering after a pin).
        private bool _suppressGridEvents;

        // Status label timer for the "Ribbon will refresh on next click." hint - cleared on next
        // user interaction or after a few seconds, whichever comes first.
        private readonly Timer _statusTimer;
        private readonly Timer _hintTimer;

        /// <summary>
        /// Constructor used by <see cref="AddIn.LanguagesButtonClicked"/>. The form owns the
        /// passed <paramref name="editable"/> instance - the caller has already cloned the live
        /// LanguageSettings via <see cref="CloneForEditing"/> so cancelling the dialog does not
        /// disturb the live state.
        /// </summary>
        public LanguageSettingsForm(AddIn addin,
                                    LanguageSettings editable,
                                    IReadOnlyList<LanguageDescriptor> registry,
                                    string jsonPath)
        {
            _addin = addin ?? throw new ArgumentNullException(nameof(addin));
            _jsonPath = jsonPath ?? throw new ArgumentNullException(nameof(jsonPath));
            _registry = registry ?? Array.Empty<LanguageDescriptor>();

            InitializeComponent();

            _statusTimer = new Timer { Interval = 4000 };
            _statusTimer.Tick += (s, e) => { _statusTimer.Stop(); lblStatus.Text = string.Empty; };

            _hintTimer = new Timer { Interval = 2500 };
            _hintTimer.Tick += (s, e) => { _hintTimer.Stop(); lblHint.Visible = false; lblHint.Text = string.Empty; };

            // Seed _allTags from registry, ordered: pinned (in pinned order), then everything else
            // alphabetically by label.
            foreach (var d in _registry)
            {
                if (d == null || string.IsNullOrEmpty(d.Tag)) continue;
                _allTags.Add(d);
            }

            if (editable != null)
            {
                _pinnedOrder.AddRange(editable.Pinned);
                foreach (var t in editable.Enabled) _enabled.Add(t);
            }

            grid.CellValueChanged += Grid_CellValueChanged;
            grid.CurrentCellDirtyStateChanged += Grid_CurrentCellDirtyStateChanged;

            RebuildGridRows();
            UpdatePinnedCountLabel();
        }

        /// <summary>
        /// Produce a deep-enough copy of <paramref name="live"/> for editing in the dialog.
        /// The dialog mutates its own copy so Cancel can discard cleanly; only Apply / OK funnel
        /// the new state through <see cref="AddIn.ReplaceLanguageSettings"/>.
        /// </summary>
        internal static LanguageSettings CloneForEditing(LanguageSettings live, IReadOnlyList<string> defaultPinned)
        {
            // When live is null, fall back to defaults so first-run / recovery still works.
            if (live == null)
            {
                return LanguageSettings.LoadOrSeed(path: null, knownTags: null,
                    defaultPinned: defaultPinned ?? new string[0]);
            }

            // Seed the clone from live.Pinned (NOT the global defaults) so any default-pinned tag
            // the user has previously disabled stays disabled. SeedDefaults writes the seed into
            // both _pinned and _enabled, so any default-pinned-but-user-disabled tag would
            // otherwise be silently resurrected in _enabled - the Enable() loop below cannot
            // un-do that, only add. Bug: a disabled "css" reappeared on every reopen.
            var clone = LanguageSettings.LoadOrSeed(path: null, knownTags: null,
                defaultPinned: live.Pinned.ToList());
            foreach (var t in live.Enabled) clone.Enable(t);
            return clone;
        }

        private void LanguageSettingsForm_Shown(object sender, EventArgs e)
        {
            // The form runs on a dedicated settings STA worker (see AddIn.EnsureSettingsWorker)
            // which never holds Windows foreground, so a bare SetForegroundWindow gets demoted to
            // a taskbar flash. Attach to the current foreground thread's input queue first - that
            // is the standard Raymond-Chen-approved way to make SetForegroundWindow stick from
            // a non-foreground thread. The TopMost flash is belt-and-braces: it forces the window
            // above OneNote on first paint without staying always-on-top afterwards.
            uint thisThread = NativeMethods.GetCurrentThreadId();
            IntPtr fg = NativeMethods.GetForegroundWindow();
            uint fgThread = fg != IntPtr.Zero ? NativeMethods.GetWindowThreadProcessId(fg, out _) : 0;

            bool attached = false;
            if (fgThread != 0 && fgThread != thisThread)
            {
                attached = NativeMethods.AttachThreadInput(thisThread, fgThread, true);
            }
            try
            {
                this.TopMost = true;
                this.TopMost = false;
                NativeMethods.SetForegroundWindow(this.Handle);
                this.Activate();
                this.Focus();
            }
            finally
            {
                if (attached) NativeMethods.AttachThreadInput(thisThread, fgThread, false);
            }
        }

        // --- Grid plumbing -----------------------------------------------------------------

        private void RebuildGridRows()
        {
            _suppressGridEvents = true;
            try
            {
                grid.Rows.Clear();
                // Pinned rows first, in order.
                foreach (var tag in _pinnedOrder)
                {
                    var d = _allTags.FirstOrDefault(x => string.Equals(x.Tag, tag, StringComparison.Ordinal));
                    if (d == null) continue;
                    AddGridRow(d, pinned: true, enabled: true);
                }

                // Then everything else, alphabetically by label.
                var rest = _allTags
                    .Where(d => !_pinnedOrder.Contains(d.Tag))
                    .OrderBy(d => d.Label ?? d.Tag, StringComparer.OrdinalIgnoreCase);
                foreach (var d in rest)
                {
                    AddGridRow(d, pinned: false, enabled: _enabled.Contains(d.Tag));
                }
            }
            finally
            {
                _suppressGridEvents = false;
            }
        }

        private void AddGridRow(LanguageDescriptor d, bool pinned, bool enabled)
        {
            int rowIdx = grid.Rows.Add(pinned, enabled, d.Label ?? d.Tag, d.Tag);
            var row = grid.Rows[rowIdx];
            row.Tag = d.Tag;
            // Pinned implies enabled: lock the E cell when P is on.
            ApplyPinEnabledLinkage(row);
        }

        private static void ApplyPinEnabledLinkage(DataGridViewRow row)
        {
            bool pinned = AsBool(row.Cells[0].Value);
            var eCell = (DataGridViewCheckBoxCell)row.Cells[1];
            if (pinned)
            {
                eCell.Value = true;
                eCell.ReadOnly = true;
                eCell.Style.BackColor = System.Drawing.SystemColors.Control;
            }
            else
            {
                eCell.ReadOnly = false;
                eCell.Style.BackColor = System.Drawing.Color.Empty;
            }
        }

        /// <summary>
        /// CommitEdit on dirty checkbox cells so CellValueChanged fires immediately rather than
        /// on focus loss. Without this the grid feels sticky.
        /// </summary>
        private void Grid_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (grid.IsCurrentCellDirty && grid.CurrentCell is DataGridViewCheckBoxCell)
            {
                grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void Grid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_suppressGridEvents) return;
            if (e.RowIndex < 0 || e.RowIndex >= grid.Rows.Count) return;
            if (e.ColumnIndex != 0 && e.ColumnIndex != 1) return;

            var row = grid.Rows[e.RowIndex];
            string tag = (string)row.Tag;
            if (string.IsNullOrEmpty(tag)) return;

            bool pinnedNow = AsBool(row.Cells[0].Value);
            bool enabledNow = AsBool(row.Cells[1].Value);

            if (e.ColumnIndex == 0) // Pinned column changed
            {
                if (pinnedNow)
                {
                    // Enforce cap. If hit, revert and flash the hint.
                    if (_pinnedOrder.Count >= MaxPinned)
                    {
                        _suppressGridEvents = true;
                        try { row.Cells[0].Value = false; }
                        finally { _suppressGridEvents = false; }
                        FlashHint("Pinned limit reached - unpin a language first");
                        return;
                    }

                    if (!_pinnedOrder.Contains(tag))
                    {
                        _pinnedOrder.Add(tag); // append; user reorders with Up/Down
                    }
                    _enabled.Add(tag);

                    _suppressGridEvents = true;
                    try
                    {
                        row.Cells[1].Value = true;
                        ApplyPinEnabledLinkage(row);
                    }
                    finally { _suppressGridEvents = false; }

                    // Move row visually to the bottom of the pinned section.
                    MoveRowToPinnedSectionBottom(row);
                }
                else
                {
                    _pinnedOrder.Remove(tag);
                    // Keep enabled. Re-enable the E cell.
                    _suppressGridEvents = true;
                    try { ApplyPinEnabledLinkage(row); }
                    finally { _suppressGridEvents = false; }

                    // Move row visually to the alphabetic position in the Available section.
                    MoveRowToAvailableSection(row);
                }
            }
            else // Enabled column changed
            {
                if (enabledNow)
                {
                    _enabled.Add(tag);
                }
                else
                {
                    // Disable removes from both. Pinned implies enabled, so unticking E on a
                    // pinned row is impossible (the cell is read-only). Belt-and-braces here.
                    _enabled.Remove(tag);
                    if (_pinnedOrder.Remove(tag))
                    {
                        _suppressGridEvents = true;
                        try { row.Cells[0].Value = false; ApplyPinEnabledLinkage(row); }
                        finally { _suppressGridEvents = false; }
                        MoveRowToAvailableSection(row);
                    }
                }
            }

            UpdatePinnedCountLabel();
            // Clear the post-Apply status hint on any user edit so it cannot leak across actions.
            lblStatus.Text = string.Empty;
            _statusTimer.Stop();
        }

        private void MoveRowToPinnedSectionBottom(DataGridViewRow row)
        {
            int targetIdx = _pinnedOrder.Count - 1; // bottom of pinned section
            if (row.Index == targetIdx) return;
            _suppressGridEvents = true;
            try
            {
                grid.Rows.Remove(row);
                grid.Rows.Insert(targetIdx, row);
            }
            finally { _suppressGridEvents = false; }
            grid.ClearSelection();
            grid.Rows[targetIdx].Selected = true;
        }

        private void MoveRowToAvailableSection(DataGridViewRow row)
        {
            // Re-render is simplest and keeps the alphabetic order in the lower section correct.
            string tag = (string)row.Tag;
            RebuildGridRows();
            // Restore selection on the moved row, if still present.
            for (int i = 0; i < grid.Rows.Count; i++)
            {
                if ((string)grid.Rows[i].Tag == tag) { grid.ClearSelection(); grid.Rows[i].Selected = true; break; }
            }
        }

        private void UpdatePinnedCountLabel()
        {
            lblPinnedCount.Text = "Pinned: " + _pinnedOrder.Count + " / " + MaxPinned;
        }

        private void FlashHint(string message)
        {
            lblHint.Text = message;
            lblHint.Visible = true;
            _hintTimer.Stop();
            _hintTimer.Start();
        }

        // --- Buttons -----------------------------------------------------------------------

        private void BtnUp_Click(object sender, EventArgs e)
        {
            if (grid.SelectedRows.Count == 0) return;
            int idx = grid.SelectedRows[0].Index;
            if (idx <= 0) return;
            if (idx >= _pinnedOrder.Count) return; // not a pinned row
            string tag = _pinnedOrder[idx];
            _pinnedOrder.RemoveAt(idx);
            _pinnedOrder.Insert(idx - 1, tag);

            RebuildGridRows();
            grid.ClearSelection();
            grid.Rows[idx - 1].Selected = true;
            UpdatePinnedCountLabel();
        }

        private void BtnDown_Click(object sender, EventArgs e)
        {
            if (grid.SelectedRows.Count == 0) return;
            int idx = grid.SelectedRows[0].Index;
            if (idx < 0) return;
            if (idx >= _pinnedOrder.Count - 1) return; // last pinned or not pinned
            string tag = _pinnedOrder[idx];
            _pinnedOrder.RemoveAt(idx);
            _pinnedOrder.Insert(idx + 1, tag);

            RebuildGridRows();
            grid.ClearSelection();
            grid.Rows[idx + 1].Selected = true;
            UpdatePinnedCountLabel();
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(this,
                "Reset to default languages? This will replace your current Pinned and Enabled lists. Apply or OK to save.",
                "NoteHighlight - Languages",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (result != DialogResult.Yes) return;

            _pinnedOrder.Clear();
            _enabled.Clear();
            foreach (var t in LanguageRegistry.DefaultPinned)
            {
                if (string.IsNullOrEmpty(t)) continue;
                if (!_pinnedOrder.Contains(t)) _pinnedOrder.Add(t);
                _enabled.Add(t);
            }
            RebuildGridRows();
            UpdatePinnedCountLabel();
            lblStatus.Text = string.Empty;
            _statusTimer.Stop();
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            if (TryPersist())
            {
                lblStatus.Text = "Ribbon will refresh on next click.";
                _statusTimer.Stop();
                _statusTimer.Start();
            }
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (TryPersist())
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        /// <summary>
        /// Build a new <see cref="LanguageSettings"/> from the in-memory model, persist it through
        /// the helper (named-mutex + atomic write), then hand it to the AddIn under its lock. On
        /// failure the live state is NOT replaced and the deferred-invalidate flag is NOT set.
        /// </summary>
        private bool TryPersist()
        {
            LanguageSettings newSettings;
            try
            {
                // Build via the LanguageSettings public API so the invariants
                // (pinned-implies-enabled, de-dupe) are enforced canonically.
                newSettings = LanguageSettings.LoadOrSeed(path: null, knownTags: null, defaultPinned: _pinnedOrder.ToList());
                foreach (var t in _enabled) newSettings.Enable(t);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Could not build language settings: " + ex.Message,
                    "NoteHighlight - Languages", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            try
            {
                SettingsHelper.SaveLanguages(newSettings, _jsonPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Could not save language settings: " + ex.Message,
                    "NoteHighlight - Languages", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Do NOT replace _languages on the AddIn, do NOT set the pending flag.
                return false;
            }

            try
            {
                _addin.ReplaceLanguageSettings(newSettings);
            }
            catch (Exception ex)
            {
                // Save succeeded but the swap failed - log only; the file is consistent.
                System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: ReplaceLanguageSettings threw after successful save. " + ex);
            }
            return true;
        }

        // --- Helpers -----------------------------------------------------------------------

        private static bool AsBool(object v)
        {
            return v is bool b && b;
        }
    }
}
