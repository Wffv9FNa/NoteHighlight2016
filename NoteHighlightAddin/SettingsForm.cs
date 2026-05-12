using GenerateHighlightContent;
using Helper;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace NoteHighlightAddin
{
    public partial class SettingsForm : Form
    {
        private const string SampleSnippet = @"// Computes the n-th Fibonacci number using memoisation.
// Demonstrates: keywords, strings, numerics, generics, comments,
// one long line that should wrap or scroll.
using System;
using System.Collections.Generic;

public static class FibDemo
{
    private static readonly Dictionary<int, long> Cache = new Dictionary<int, long>();

    public static long Fib(int n)
    {
        if (n < 0) throw new ArgumentOutOfRangeException(""n"", ""Fibonacci is undefined for negative indices."");
        if (n < 2) return n;
        if (Cache.TryGetValue(n, out long cached)) return cached;
        long value = Fib(n - 1) + Fib(n - 2);
        Cache[n] = value;
        return value;
    }

    public static void Main() { Console.WriteLine(""Fib(40) = "" + Fib(40)); }
}";

        /// <summary>
        /// Set while the splitter is being positioned programmatically (initial load,
        /// resize-driven re-apply) so the SplitterMoved handler does not write the
        /// transient value back to user.config.
        /// </summary>
        private bool _suppressSplitterPersist;

        public SettingsForm()
        {
            InitializeComponent();

            fontDialog1.Font = new Font(NoteHighlightForm.Properties.Settings.Default.Font, NoteHighlightForm.Properties.Settings.Default.FontSize);
            btnFont.Text = "Font:" + fontDialog1.Font.Name + ", Size:" + fontDialog1.Font.Size;
            btnFont.Font = fontDialog1.Font;
            cbShowTableBorder.Checked = NoteHighlightForm.Properties.Settings.Default.ShowTableBorder;

            this.splitContainer.SplitterMoved += SplitContainer_SplitterMoved;
        }

        /// <summary>
        /// Applies the persisted splitter percentage to the SplitContainer. Honours
        /// Panel1MinSize and Panel2MinSize so the SplitContainer does not throw.
        /// </summary>
        private void ApplySavedSplitterDistance()
        {
            if (this.splitContainer == null) return;
            int width = this.splitContainer.Width;
            if (width <= 0) return;

            int percent = NoteHighlightForm.Properties.Settings.Default.SettingsFormPreviewSplitter;
            if (percent <= 0 || percent >= 100) percent = 60;

            int desired = (int)Math.Round(width * (percent / 100.0));

            int min = this.splitContainer.Panel1MinSize;
            int max = width - this.splitContainer.Panel2MinSize - this.splitContainer.SplitterWidth;
            if (max < min) return;

            int clamped = Math.Max(min, Math.Min(max, desired));

            _suppressSplitterPersist = true;
            try
            {
                this.splitContainer.SplitterDistance = clamped;
            }
            finally
            {
                _suppressSplitterPersist = false;
            }
        }

        private void SplitContainer_SplitterMoved(object sender, SplitterEventArgs e)
        {
            if (_suppressSplitterPersist) return;
            if (!this.IsHandleCreated) return;

            int width = this.splitContainer.Width;
            if (width <= 0) return;

            int percent = (int)Math.Round(this.splitContainer.SplitterDistance * 100.0 / width);
            if (percent < 1) percent = 1;
            if (percent > 99) percent = 99;

            if (NoteHighlightForm.Properties.Settings.Default.SettingsFormPreviewSplitter == percent) return;

            NoteHighlightForm.Properties.Settings.Default.SettingsFormPreviewSplitter = percent;
            SettingsHelper.SafeSave();
        }

        private void BtnFont_Click(object sender, EventArgs e)
        {
            fontDialog1.Font = new Font(NoteHighlightForm.Properties.Settings.Default.Font, NoteHighlightForm.Properties.Settings.Default.FontSize);
            if (fontDialog1.ShowDialog() == DialogResult.OK)
            {
                btnFont.Text = "Font:"+fontDialog1.Font.Name + ", Size:" + fontDialog1.Font.Size;
                btnFont.Font = fontDialog1.Font;

                NoteHighlightForm.Properties.Settings.Default.Font = fontDialog1.Font.Name;
                NoteHighlightForm.Properties.Settings.Default.FontSize = (int)Math.Round(fontDialog1.Font.Size);

                SettingsHelper.SafeSave();
                SchedulePreview();
            }


        }

        private void ChShowTableBorder_CheckedChanged(object sender, EventArgs e)
        {
            NoteHighlightForm.Properties.Settings.Default.ShowTableBorder = cbShowTableBorder.Checked;

            SettingsHelper.SafeSave();
            SchedulePreview();
        }

        private void SchedulePreview()
        {
            if (_previewPane == null) return;

            var defaults = NoteHighlightForm.Properties.Settings.Default;
            var parameters = new HighLightParameter()
            {
                FileName = "settings-preview",
                Content = SampleSnippet,
                CodeType = "cs",
                HighLightStyle = ResolveThemeName(defaults.HighLightStyle),
                ShowLineNumber = defaults.ShowLineNumber,
                HighlightColor = defaults.BackgroundColor,
                Font = fontDialog1.Font.Name,
                FontSize = (int)Math.Round(fontDialog1.Font.Size)
            };

            _previewPane.Render(parameters, defaults.DarkMode);
        }

        private static string ResolveThemeName(int themeIndex)
        {
            try
            {
                var section = (new GenerateHighLight(AddIn.GetAddinDirectory())).Config;
                var assemblyLocation = typeof(SettingsForm).Assembly.Location;
                if (string.IsNullOrEmpty(assemblyLocation))
                    assemblyLocation = new Uri(typeof(SettingsForm).Assembly.CodeBase).LocalPath;
                var themesDir = Path.Combine(ProcessHelper.GetDirectoryFromPath(assemblyLocation), section.FolderName, section.ThemeFolder);
                var themes = Directory.GetFiles(themesDir, "*.theme");
                if (themes.Length == 0) return string.Empty;
                int idx = themeIndex >= 0 && themeIndex < themes.Length ? themeIndex : 0;
                return Path.GetFileNameWithoutExtension(themes[idx]);
            }
            catch
            {
                return string.Empty;
            }
        }

        private void SettingsForm_Shown(object sender, EventArgs e)
        {
            // This is necessary in order for SetForegroundWindow to work consistently
            this.WindowState = FormWindowState.Minimized;
            this.WindowState = FormWindowState.Normal;

            NativeMethods.SetForegroundWindow(this.Handle);

            ApplySavedPreviewVisibility();
            if (!this.splitContainer.Panel2Collapsed)
            {
                ApplySavedSplitterDistance();
            }

            this.BeginInvoke(new Action(SchedulePreview));
        }

        /// <summary>
        /// Reads the persisted preview-visibility flag and applies it. Must be called before
        /// ApplySavedSplitterDistance: when Panel2Collapsed is true the splitter distance is
        /// invalid (the SplitContainer will throw if its constraints can't be honoured).
        /// </summary>
        private void ApplySavedPreviewVisibility()
        {
            bool visible = NoteHighlightForm.Properties.Settings.Default.SettingsFormPreviewVisible;
            this.splitContainer.Panel2Collapsed = !visible;
            UpdateTogglePreviewButtonText(visible);
        }

        private void UpdateTogglePreviewButtonText(bool previewVisible)
        {
            this.btnTogglePreview.Text = previewVisible ? "Hide preview" : "Show preview";
        }

        private void btnTogglePreview_Click(object sender, EventArgs e)
        {
            bool nowVisible = this.splitContainer.Panel2Collapsed; // collapsed -> will become visible
            this.splitContainer.Panel2Collapsed = !nowVisible;
            UpdateTogglePreviewButtonText(nowVisible);

            NoteHighlightForm.Properties.Settings.Default.SettingsFormPreviewVisible = nowVisible;
            SettingsHelper.SafeSave();

            if (nowVisible)
            {
                ApplySavedSplitterDistance();
                SchedulePreview();
            }
        }

        private void SettingsForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _previewPane?.Dispose();
        }
    }
}
