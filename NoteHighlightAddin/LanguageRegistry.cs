/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace NoteHighlightAddin
{
    /// <summary>
    /// Read-only descriptor for a single language button declared in ribbon.xml.
    /// </summary>
    public sealed class LanguageDescriptor
    {
        public string ButtonId  { get; }
        public string Tag       { get; }
        public string Label     { get; }
        public string Screentip { get; }
        public string Image     { get; }
        public string Keytip    { get; }

        public LanguageDescriptor(string buttonId, string tag, string label, string screentip, string image, string keytip)
        {
            ButtonId  = buttonId;
            Tag       = tag;
            Label     = label;
            Screentip = screentip;
            Image     = image;
            Keytip    = keytip;
        }
    }

    /// <summary>
    /// One-shot, fail-soft parser for ribbon.xml. The canonical list of languages
    /// the add-in offers is the set of buttons under the Language group(s) in
    /// ribbon.xml; this registry exposes that list at runtime for callers that
    /// need to map a tag to a descriptor (e.g. LanguageSettings invariants,
    /// the dynamic-menu populator in Phase 2).
    ///
    /// <para>
    /// Failure handling (plan section 4.3): Initialise never throws. On any
    /// parse or I/O failure we log loudly via Trace and leave <see cref="All"/>
    /// empty so callers can serve a minimal recovery UI. The caller is
    /// expected to check <see cref="All"/>.Count and act accordingly.
    /// </para>
    /// </summary>
    public static class LanguageRegistry
    {
        // The default-pinned tag list mirrors what is currently visible="true" in
        // ribbon.xml so the user-visible ribbon does not change after upgrade.
        // The "section 3.2" defaults (cs/sql/py/js/...) are Phase 3 territory;
        // Phase 1 deliberately preserves today's appearance.
        private static readonly string[] PhaseOneDefaultPinned = new[]
        {
            "cs", "sql", "css", "js", "html", "xml", "java",
            "php", "perl", "py", "ruby", "c", "ps1",
        };

        private static readonly object SyncRoot = new object();
        private static LanguageDescriptor[] _all = Array.Empty<LanguageDescriptor>();
        private static Dictionary<string, LanguageDescriptor> _byTag =
            new Dictionary<string, LanguageDescriptor>(StringComparer.Ordinal);
        private static Dictionary<string, LanguageDescriptor> _byButtonId =
            new Dictionary<string, LanguageDescriptor>(StringComparer.Ordinal);
        private static int _initialised; // 0 = no, 1 = yes (Interlocked)

        /// <summary>True once <see cref="Initialise(string)"/> has run (successfully or not).</summary>
        public static bool IsInitialised => Volatile.Read(ref _initialised) == 1;

        /// <summary>All language buttons declared in ribbon.xml. Empty if parsing failed.</summary>
        public static IReadOnlyList<LanguageDescriptor> All
        {
            get { lock (SyncRoot) { return _all; } }
        }

        /// <summary>
        /// Phase 1 seed list - the set of <c>visible="true"</c> tags in today's ribbon.xml.
        /// Stable across runs; not derived from the parsed file.
        /// </summary>
        public static IReadOnlyList<string> DefaultPinned => PhaseOneDefaultPinned;

        /// <summary>
        /// Look up a descriptor by <c>tag</c> attribute (e.g. "cs"). Returns null
        /// if not found or if the registry never initialised successfully.
        /// </summary>
        public static LanguageDescriptor ByTag(string tag)
        {
            if (string.IsNullOrEmpty(tag)) return null;
            lock (SyncRoot)
            {
                _byTag.TryGetValue(tag, out var d);
                return d;
            }
        }

        /// <summary>Look up a descriptor by button id (e.g. "buttonCSharp").</summary>
        public static LanguageDescriptor ByButtonId(string buttonId)
        {
            if (string.IsNullOrEmpty(buttonId)) return null;
            lock (SyncRoot)
            {
                _byButtonId.TryGetValue(buttonId, out var d);
                return d;
            }
        }

        /// <summary>
        /// The on-disk filename used by the dev / diagnostic hot-edit fallback.
        /// Renamed from <c>ribbon.xml</c> to <c>ribbon.override.xml</c> as part
        /// of bug 13.1 followup (2026-05-13 evening): end-user installs must
        /// never have this file present, so we picked a name the MSI does not
        /// ship. See <see cref="LegacyOnDiskRibbonFileName"/> for the legacy
        /// safety-net path.
        /// </summary>
        public const string OnDiskRibbonFileName = "ribbon.override.xml";

        /// <summary>
        /// Legacy hot-edit filename. Older installs (&lt;= 3.8) wrote
        /// <c>ribbon.xml</c> to the addin directory via the MSI content list,
        /// and the component was marked <c>Permanent=TRUE</c> so uninstall did
        /// not remove it. If such a file is found next to the DLL we IGNORE it
        /// (Trace warning) and use the embedded copy. The file is now harmless
        /// dead data; users may delete it manually but the add-in will work
        /// correctly either way.
        /// </summary>
        internal const string LegacyOnDiskRibbonFileName = "ribbon.xml";

        /// <summary>
        /// Parse the embedded ribbon.xml resource, preferring an on-disk
        /// <c>ribbon.override.xml</c> next to the DLL if present (dev /
        /// diagnostic hot-edit). Never throws. Equivalent to
        /// <see cref="Initialise(string)"/> with the addin directory resolved
        /// from <see cref="AddIn.GetAddinDirectory"/>. See bug 13.1
        /// (.local/docs/bugs/msi-ribbon-stale-on-upgrade.md) for why the
        /// canonical source is the embedded copy rather than an on-disk file.
        /// </summary>
        public static void Initialise()
        {
            Initialise(AddIn.GetAddinDirectory());
        }

        /// <summary>
        /// Parse ribbon.xml. Idempotent (subsequent calls are a no-op). Never
        /// throws - on any failure the registry is left empty and a Trace
        /// warning is emitted.
        ///
        /// <para>
        /// Source order (post bug 13.1 followup): if
        /// <paramref name="addinDirectory"/> contains a
        /// <c>ribbon.override.xml</c> file on disk, parse that (dev /
        /// diagnostic hot-edit convenience). Otherwise fall back to the
        /// embedded <c>Properties.Resources.ribbon</c> copy. A legacy
        /// <c>ribbon.xml</c> next to the DLL is deliberately IGNORED (with a
        /// Trace warning logging the resolved path) - older MSIs marked the
        /// component <c>Permanent=TRUE</c>, so the file persists across
        /// uninstall and would otherwise beat the embedded copy with stale
        /// content.
        /// </para>
        ///
        /// <para>
        /// The path overload remains for unit-test fixtures that write a
        /// ribbon.override.xml to a temp directory and want it picked up
        /// deterministically; production callers should prefer
        /// <see cref="Initialise()"/>.
        /// </para>
        /// </summary>
        public static void Initialise(string addinDirectory)
        {
            // Idempotent: only the first caller does real work.
            if (Interlocked.CompareExchange(ref _initialised, 1, 0) != 0)
                return;

            string overridePath = null;
            try
            {
                XDocument doc = null;
                string sourceDescription = null;

                if (!string.IsNullOrEmpty(addinDirectory))
                {
                    // 1a. Legacy-file safety net. If an old ribbon.xml is
                    //     sitting next to the DLL (left behind by a pre-3.9
                    //     install whose Permanent=TRUE component was never
                    //     uninstalled, or a hand-edit by a power user), log
                    //     and ignore it - the embedded copy is canonical and
                    //     the legacy file is almost certainly stale. Memory
                    //     rule "feedback_com_addin_path_traps.md": validate
                    //     File.Exists and log the resolved path.
                    var legacyPath = Path.Combine(addinDirectory, LegacyOnDiskRibbonFileName);
                    if (File.Exists(legacyPath))
                    {
                        Trace.TraceWarning("NoteHighlight2016: ignoring legacy on-disk '" + legacyPath + "' - the embedded ribbon resource is canonical. You may delete this file manually; the add-in no longer reads it.");
                    }

                    // 1b. Hot-edit / diagnostic path: a ribbon.override.xml
                    //     file next to the DLL. End-user installs never ship
                    //     this file - it only exists if a developer placed it
                    //     there intentionally.
                    overridePath = Path.Combine(addinDirectory, OnDiskRibbonFileName);
                    if (File.Exists(overridePath))
                    {
                        try
                        {
                            doc = XDocument.Load(overridePath);
                            sourceDescription = overridePath;
                            Trace.TraceInformation("NoteHighlight2016: using on-disk ribbon override '" + overridePath + "'.");
                        }
                        catch (Exception ex)
                        {
                            Trace.TraceWarning("NoteHighlight2016: failed to parse on-disk '" + overridePath + "'; falling back to embedded resource. " + ex.GetType().Name + ": " + ex.Message);
                            doc = null;
                        }
                    }
                }

                // 2. Embedded canonical copy.
                if (doc == null)
                {
                    string embedded;
                    try
                    {
                        embedded = Properties.Resources.ribbon;
                    }
                    catch (Exception ex)
                    {
                        Trace.TraceError("NoteHighlight2016: failed to read embedded ribbon resource; " + ex.GetType().Name + ": " + ex.Message);
                        return;
                    }

                    if (string.IsNullOrEmpty(embedded))
                    {
                        Trace.TraceError("NoteHighlight2016: embedded ribbon resource is empty and no on-disk ribbon.override.xml found (looked at '" + (overridePath ?? "<no addin directory>") + "'); LanguageRegistry left empty.");
                        return;
                    }

                    try
                    {
                        doc = XDocument.Parse(embedded);
                        sourceDescription = "<embedded Resources.ribbon>";
                    }
                    catch (Exception ex)
                    {
                        Trace.TraceError("NoteHighlight2016: failed to parse embedded ribbon resource; " + ex.GetType().Name + ": " + ex.Message);
                        return;
                    }
                }

                var root = doc.Root;
                if (root == null)
                {
                    Trace.TraceError("NoteHighlight2016: ribbon.xml (" + (sourceDescription ?? "<unknown source>") + ") has no root element.");
                    return;
                }

                // Walk language groups by id. Today there is one ("groupLanguage")
                // but tolerate any group whose id starts with "groupLanguage".
                var languageGroups = root.Descendants()
                    .Where(e => e.Name.LocalName == "group"
                                && (string)e.Attribute("id") != null
                                && ((string)e.Attribute("id")).StartsWith("groupLanguage", StringComparison.Ordinal))
                    .ToList();

                var descriptors = new List<LanguageDescriptor>();
                foreach (var btn in languageGroups.SelectMany(g => g.Descendants().Where(e => e.Name.LocalName == "button")))
                {
                    var id        = (string)btn.Attribute("id");
                    var tag       = (string)btn.Attribute("tag");
                    var label     = (string)btn.Attribute("label");
                    var screentip = (string)btn.Attribute("screentip");
                    var image     = (string)btn.Attribute("image");
                    var keytip    = (string)btn.Attribute("keytip");

                    if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(tag))
                    {
                        // Slot buttons (slotLangNN) legitimately have no tag - they are not languages.
                        if (id == null || !id.StartsWith("slotLang", StringComparison.Ordinal))
                            Trace.TraceWarning("NoteHighlight2016: skipping ribbon button without id/tag (id='" + (id ?? "") + "', tag='" + (tag ?? "") + "').");
                        continue;
                    }

                    descriptors.Add(new LanguageDescriptor(id, tag, label ?? id, screentip, image, keytip));
                }

                lock (SyncRoot)
                {
                    _all = descriptors.ToArray();
                    _byTag = descriptors
                        .GroupBy(d => d.Tag, StringComparer.Ordinal)
                        .ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal);
                    _byButtonId = descriptors
                        .GroupBy(d => d.ButtonId, StringComparer.Ordinal)
                        .ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal);
                }
            }
            catch (Exception ex)
            {
                // Defensive: never let an unanticipated exception escape into
                // OnConnection. Log with the resolved path so the user can see
                // where we looked.
                Trace.TraceError("NoteHighlight2016: unexpected error initialising LanguageRegistry from '" + (overridePath ?? addinDirectory ?? "<null>") + "'; " + ex);
            }
        }

        /// <summary>
        /// Test-only reset hook. Internal so production code cannot accidentally
        /// double-initialise. Used by unit tests that exercise the parse path
        /// with multiple ribbon.xml fixtures in one process.
        /// </summary>
        internal static void ResetForTests()
        {
            lock (SyncRoot)
            {
                _all = Array.Empty<LanguageDescriptor>();
                _byTag = new Dictionary<string, LanguageDescriptor>(StringComparer.Ordinal);
                _byButtonId = new Dictionary<string, LanguageDescriptor>(StringComparer.Ordinal);
            }
            Volatile.Write(ref _initialised, 0);
        }
    }
}
