/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace NoteHighlightAddin
{
    /// <summary>
    /// Per-user language picker state, persisted as JSON to
    /// <c>%APPDATA%\NoteHighlight2016\languages.json</c>. See plan section 3.
    ///
    /// <para>
    /// Phase 1 lands the type and the persistence pipeline; the dynamic menu
    /// and the settings dialog ride on top in later phases. The public
    /// surface intentionally hides the underlying lists behind
    /// <see cref="IsPinned"/> / <see cref="IsEnabled"/> /
    /// <see cref="IsVisibleAsPinnedButton"/> so consumers cannot mutate state
    /// accidentally.
    /// </para>
    /// </summary>
    public sealed class LanguageSettings
    {
        // Suffix derives from the add-in's COM GUID 4C6B0362-F139-417F-9661-3663C268B9E9.
        // Global\ scopes the mutex to the machine; per-user scoping is unnecessary
        // because the file path is already per-user.
        private const string MutexName = @"Global\NoteHighlight2016-4C6B0362-languages";

        private const int CurrentSchemaVersion = 1;

        private readonly List<string> _pinned;
        private readonly HashSet<string> _enabled;

        // Internal ctor; consumers go through LoadOrSeed.
        private LanguageSettings(IEnumerable<string> pinned, IEnumerable<string> enabled)
        {
            _pinned  = pinned.ToList();
            _enabled = new HashSet<string>(enabled, StringComparer.Ordinal);

            // Invariant: pinned implies enabled.
            foreach (var tag in _pinned)
            {
                _enabled.Add(tag);
            }
        }

        public int SchemaVersion => CurrentSchemaVersion;

        /// <summary>Pinned tags in user-chosen ribbon order.</summary>
        public IReadOnlyList<string> Pinned => _pinned;

        /// <summary>All enabled tags (pinned tags are members; full set, not the difference).</summary>
        public IReadOnlyCollection<string> Enabled => _enabled;

        public bool IsPinned(string tag) => !string.IsNullOrEmpty(tag) && _pinned.Contains(tag);

        public bool IsEnabled(string tag) => !string.IsNullOrEmpty(tag) && _enabled.Contains(tag);

        /// <summary>Returns true if the language should appear as a top-level pinned ribbon button.</summary>
        public bool IsVisibleAsPinnedButton(string tag) => IsPinned(tag);

        /// <summary>Returns true if the language should appear in the More languages... dynamic menu.</summary>
        public bool IsVisibleInMoreMenu(string tag) => IsEnabled(tag) && !IsPinned(tag);

        /// <summary>
        /// Pin <paramref name="tag"/> at <paramref name="position"/>. If already
        /// pinned the tag is moved. Pinned implies enabled; enabled membership
        /// is added if absent. Position is clamped to <c>[0, count]</c>.
        /// </summary>
        public void Pin(string tag, int position)
        {
            if (string.IsNullOrEmpty(tag)) throw new ArgumentNullException(nameof(tag));
            _pinned.Remove(tag);
            if (position < 0) position = 0;
            if (position > _pinned.Count) position = _pinned.Count;
            _pinned.Insert(position, tag);
            _enabled.Add(tag);
        }

        /// <summary>Unpin a tag. The tag remains enabled.</summary>
        public void Unpin(string tag)
        {
            if (string.IsNullOrEmpty(tag)) return;
            _pinned.Remove(tag);
        }

        public void Enable(string tag)
        {
            if (string.IsNullOrEmpty(tag)) return;
            _enabled.Add(tag);
        }

        /// <summary>Disable a tag. Removes it from both pinned and enabled sets.</summary>
        public void Disable(string tag)
        {
            if (string.IsNullOrEmpty(tag)) return;
            _pinned.Remove(tag);
            _enabled.Remove(tag);
        }

        /// <summary>
        /// Replace the pinned order with <paramref name="newOrder"/>. Tags not in
        /// <paramref name="newOrder"/> are dropped from pinned (but kept in enabled).
        /// </summary>
        public void Reorder(IList<string> newOrder)
        {
            if (newOrder == null) throw new ArgumentNullException(nameof(newOrder));
            _pinned.Clear();
            foreach (var tag in newOrder)
            {
                if (string.IsNullOrEmpty(tag)) continue;
                if (!_pinned.Contains(tag))
                {
                    _pinned.Add(tag);
                    _enabled.Add(tag);
                }
            }
        }

        /// <summary>
        /// Load settings from <paramref name="path"/>, or seed defaults if the
        /// file is missing, corrupt, or references unknown tags. Never throws -
        /// any failure falls back to defaults with a Trace warning.
        /// </summary>
        /// <param name="path">Final file path (typically %APPDATA%\NoteHighlight2016\languages.json).</param>
        /// <param name="knownTags">The current set of tags declared in ribbon.xml; entries not in this set are dropped silently. Pass an empty list to disable filtering.</param>
        /// <param name="defaultPinned">Defaults applied when the file does not yet exist or is unrecoverable.</param>
        public static LanguageSettings LoadOrSeed(string path,
                                                  IReadOnlyList<string> knownTags,
                                                  IReadOnlyList<string> defaultPinned)
        {
            if (defaultPinned == null) throw new ArgumentNullException(nameof(defaultPinned));
            var knownTagSet = knownTags != null && knownTags.Count > 0
                ? new HashSet<string>(knownTags, StringComparer.Ordinal)
                : null;

            // Missing file: seed defaults, write through, return.
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                var seeded = SeedDefaults(defaultPinned, knownTagSet);
                TrySeedWrite(seeded, path);
                return seeded;
            }

            string text = null;
            try
            {
                text = File.ReadAllText(path, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                Trace.TraceWarning("NoteHighlight2016: could not read '" + path + "'; using defaults. " + ex.Message);
                return SeedDefaults(defaultPinned, knownTagSet);
            }

            Dto dto = null;
            try
            {
                dto = new JavaScriptSerializer().Deserialize<Dto>(text);
            }
            catch (Exception ex)
            {
                Trace.TraceWarning("NoteHighlight2016: malformed JSON in '" + path + "'; using defaults. " + ex.Message);
                TryWriteBakAndReturnDefaults(path, defaultPinned, knownTagSet, out var seeded);
                return seeded;
            }

            if (dto == null || dto.schemaVersion != CurrentSchemaVersion)
            {
                Trace.TraceWarning("NoteHighlight2016: '" + path + "' has unsupported schemaVersion (" +
                                   (dto?.schemaVersion ?? -1) + "); using defaults.");
                TryWriteBakAndReturnDefaults(path, defaultPinned, knownTagSet, out var seeded);
                return seeded;
            }

            // Apply invariants: filter unknown tags, dedupe, force pinned-implies-enabled.
            var pinned = NormaliseList(dto.pinned, knownTagSet, path, "pinned");
            var enabled = NormaliseList(dto.enabled, knownTagSet, path, "enabled");
            return new LanguageSettings(pinned, enabled);
        }

        /// <summary>
        /// Persist this instance to <paramref name="finalPath"/>. Atomic via
        /// <see cref="File.Replace(string, string, string, bool)"/> (or
        /// <see cref="File.Move(string, string)"/> on first write), gated by a
        /// machine-wide named mutex and an in-process lock. See plan section 3.3.
        /// </summary>
        public void SaveAtomically(string finalPath)
        {
            if (string.IsNullOrEmpty(finalPath)) throw new ArgumentNullException(nameof(finalPath));

            var dir = Path.GetDirectoryName(finalPath);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir); // no-op if it already exists
            }

            var tempPath = finalPath + ".tmp";
            var json = SerializeToJson(this);

            bool createdNew;
            using (var mtx = new Mutex(false, MutexName, out createdNew))
            {
                bool acquired = false;
                try
                {
                    acquired = mtx.WaitOne(TimeSpan.FromSeconds(5));
                }
                catch (AbandonedMutexException)
                {
                    // The previous owner crashed without releasing. We now own
                    // the mutex; proceed with the write.
                    acquired = true;
                }

                if (!acquired)
                {
                    Trace.TraceWarning("NoteHighlight2016: languages.json save - mutex timeout for '" + finalPath + "'.");
                    return;
                }

                try
                {
                    lock (SettingsHelper.SaveLock)
                    {
                        File.WriteAllText(tempPath, json, new UTF8Encoding(false));

                        int attempts = 5;
                        while (true)
                        {
                            try
                            {
                                if (File.Exists(finalPath))
                                {
                                    File.Replace(tempPath, finalPath, null, ignoreMetadataErrors: true);
                                }
                                else
                                {
                                    File.Move(tempPath, finalPath);
                                }
                                return;
                            }
                            catch (IOException) when (--attempts > 0)
                            {
                                Thread.Sleep(50); // bounded retry for AV / search-indexer races
                            }
                        }
                    }
                }
                finally
                {
                    try { mtx.ReleaseMutex(); } catch { /* never throw out of save */ }
                }
            }
        }

        // --- Private helpers ----------------------------------------------------

        private static LanguageSettings SeedDefaults(IReadOnlyList<string> defaultPinned, HashSet<string> knownTagSet)
        {
            var pinned = new List<string>();
            foreach (var tag in defaultPinned)
            {
                if (string.IsNullOrEmpty(tag)) continue;
                if (pinned.Contains(tag)) continue;
                if (knownTagSet != null && !knownTagSet.Contains(tag))
                {
                    Trace.TraceWarning("NoteHighlight2016: default-pinned tag '" + tag + "' not declared in ribbon.xml; dropping from seed.");
                    continue;
                }
                pinned.Add(tag);
            }
            return new LanguageSettings(pinned, pinned);
        }

        private static void TrySeedWrite(LanguageSettings seeded, string path)
        {
            if (string.IsNullOrEmpty(path)) return;
            try
            {
                seeded.SaveAtomically(path);
            }
            catch (Exception ex)
            {
                Trace.TraceWarning("NoteHighlight2016: could not write seed languages.json to '" + path + "'; in-memory defaults used. " + ex.Message);
            }
        }

        private static void TryWriteBakAndReturnDefaults(string path,
                                                         IReadOnlyList<string> defaultPinned,
                                                         HashSet<string> knownTagSet,
                                                         out LanguageSettings seeded)
        {
            seeded = SeedDefaults(defaultPinned, knownTagSet);
            try
            {
                if (File.Exists(path))
                {
                    var bak = path + ".bak";
                    try { if (File.Exists(bak)) File.Delete(bak); } catch { /* best-effort */ }
                    File.Move(path, bak);
                    Trace.TraceWarning("NoteHighlight2016: moved corrupt '" + path + "' to '" + bak + "'.");
                }
            }
            catch (Exception ex)
            {
                Trace.TraceWarning("NoteHighlight2016: could not back up '" + path + "'; continuing with defaults. " + ex.Message);
            }
            TrySeedWrite(seeded, path);
        }

        private static List<string> NormaliseList(string[] raw, HashSet<string> knownTagSet, string path, string label)
        {
            var result = new List<string>();
            if (raw == null) return result;
            foreach (var t in raw)
            {
                if (string.IsNullOrEmpty(t)) continue;
                if (result.Contains(t)) continue; // de-dupe (first wins)
                if (knownTagSet != null && !knownTagSet.Contains(t))
                {
                    Trace.TraceWarning("NoteHighlight2016: dropping unknown tag '" + t + "' from '" + label + "' in '" + path + "'.");
                    continue;
                }
                result.Add(t);
            }
            return result;
        }

        private static string SerializeToJson(LanguageSettings s)
        {
            var dto = new Dto
            {
                schemaVersion = CurrentSchemaVersion,
                pinned  = s._pinned.ToArray(),
                enabled = s._enabled.OrderBy(t => t, StringComparer.Ordinal).ToArray(),
            };
            return new JavaScriptSerializer().Serialize(dto);
        }

        // DTO. Lower-case field names match the JSON wire format; JavaScriptSerializer
        // is case-sensitive on field bindings.
#pragma warning disable IDE1006, CS0649 // intentional lower-case wire names; fields populated by reflection
        internal sealed class Dto
        {
            public int schemaVersion;
            public string[] pinned;
            public string[] enabled;
        }
#pragma warning restore IDE1006, CS0649
    }
}
