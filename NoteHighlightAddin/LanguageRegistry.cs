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
        /// Parse ribbon.xml from <paramref name="addinDirectory"/>. Idempotent
        /// (subsequent calls are a no-op). Never throws - on any failure the
        /// registry is left empty and a Trace warning is emitted.
        /// </summary>
        public static void Initialise(string addinDirectory)
        {
            // Idempotent: only the first caller does real work.
            if (Interlocked.CompareExchange(ref _initialised, 1, 0) != 0)
                return;

            string ribbonPath = null;
            try
            {
                if (string.IsNullOrEmpty(addinDirectory))
                {
                    Trace.TraceError("NoteHighlight2016: LanguageRegistry.Initialise called with empty addinDirectory; leaving registry empty.");
                    return;
                }

                ribbonPath = Path.Combine(addinDirectory, "ribbon.xml");

                // Memory-rule "feedback_com_addin_path_traps.md": validate
                // File.Exists and fail loudly with the resolved path so callers
                // can debug COM-activation path drift.
                if (!File.Exists(ribbonPath))
                {
                    Trace.TraceError("NoteHighlight2016: ribbon.xml not found at '" + ribbonPath + "'; LanguageRegistry left empty.");
                    return;
                }

                XDocument doc;
                try
                {
                    doc = XDocument.Load(ribbonPath);
                }
                catch (Exception ex)
                {
                    Trace.TraceError("NoteHighlight2016: failed to parse '" + ribbonPath + "'; " + ex.GetType().Name + ": " + ex.Message);
                    return;
                }

                var root = doc.Root;
                if (root == null)
                {
                    Trace.TraceError("NoteHighlight2016: ribbon.xml at '" + ribbonPath + "' has no root element.");
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
                Trace.TraceError("NoteHighlight2016: unexpected error initialising LanguageRegistry from '" + (ribbonPath ?? addinDirectory ?? "<null>") + "'; " + ex);
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
