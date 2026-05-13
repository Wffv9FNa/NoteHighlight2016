/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NoteHighlightAddin;

namespace UnitTesting
{
    /// <summary>
    /// Unit tests for <see cref="LanguageSettings"/>. Plan section 9.1.
    /// Each test uses a unique temp directory so saves do not collide across
    /// the suite, and the directory is removed in TestCleanup.
    /// </summary>
    [TestClass]
    public class LanguageSettingsTests
    {
        private string _tempDir;
        private string _path;

        private static readonly string[] DefaultsPhase1 = new[]
        {
            "cs", "sql", "css", "js", "html", "xml", "java",
            "php", "perl", "py", "ruby", "c", "ps1",
        };

        // A "known tags" set including the defaults plus a couple of extras so
        // tests can assert that legitimately-extra tags survive normalisation.
        private static readonly string[] KnownTags = new[]
        {
            "cs", "sql", "css", "js", "html", "xml", "java",
            "php", "perl", "py", "ruby", "c", "ps1",
            "go", "ts", "md", "json",
        };

        [TestInitialize]
        public void SetUp()
        {
            _tempDir = Path.Combine(Path.GetTempPath(), "nh-tests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
            _path = Path.Combine(_tempDir, "languages.json");
        }

        [TestCleanup]
        public void TearDown()
        {
            try { if (Directory.Exists(_tempDir)) Directory.Delete(_tempDir, true); } catch { /* best-effort */ }
        }

        // --- LoadOrSeed --------------------------------------------------------

        [TestMethod]
        public void LoadOrSeed_MissingFile_SeedsAndPersists()
        {
            Assert.IsFalse(File.Exists(_path));

            var s = LanguageSettings.LoadOrSeed(_path, KnownTags, DefaultsPhase1);

            Assert.IsTrue(File.Exists(_path), "Seed path should have written the file.");
            CollectionAssert.AreEqual(DefaultsPhase1, s.Pinned.ToArray());
            foreach (var t in DefaultsPhase1)
            {
                Assert.IsTrue(s.IsEnabled(t), "Seeded tag '" + t + "' should be enabled.");
            }
        }

        [TestMethod]
        public void LoadOrSeed_MissingFile_NoKnownTagsFilter_StillSeeds()
        {
            var s = LanguageSettings.LoadOrSeed(_path, null, DefaultsPhase1);

            Assert.AreEqual(DefaultsPhase1.Length, s.Pinned.Count);
        }

        [TestMethod]
        public void LoadOrSeed_CorruptJson_FallsBackAndWritesBak()
        {
            File.WriteAllText(_path, "this is not json {{{", new UTF8Encoding(false));

            var s = LanguageSettings.LoadOrSeed(_path, KnownTags, DefaultsPhase1);

            Assert.AreEqual(DefaultsPhase1.Length, s.Pinned.Count, "Corrupt input must fall back to defaults.");
            Assert.IsTrue(File.Exists(_path + ".bak"), "Corrupt original should be moved to .bak.");
        }

        [TestMethod]
        public void LoadOrSeed_UnknownTag_DroppedSilently()
        {
            File.WriteAllText(_path,
                "{\"schemaVersion\":1,\"pinned\":[\"cs\",\"not-a-real-tag\",\"py\"],\"enabled\":[\"cs\",\"not-a-real-tag\",\"py\"]}",
                new UTF8Encoding(false));

            var s = LanguageSettings.LoadOrSeed(_path, KnownTags, DefaultsPhase1);

            Assert.IsTrue(s.IsPinned("cs"));
            Assert.IsTrue(s.IsPinned("py"));
            Assert.IsFalse(s.IsPinned("not-a-real-tag"));
            Assert.IsFalse(s.IsEnabled("not-a-real-tag"));
        }

        [TestMethod]
        public void LoadOrSeed_UnknownSchemaVersion_FallsBackToDefaults()
        {
            File.WriteAllText(_path,
                "{\"schemaVersion\":999,\"pinned\":[\"cs\"],\"enabled\":[\"cs\"]}",
                new UTF8Encoding(false));

            var s = LanguageSettings.LoadOrSeed(_path, KnownTags, DefaultsPhase1);

            Assert.AreEqual(DefaultsPhase1.Length, s.Pinned.Count);
        }

        [TestMethod]
        public void LoadOrSeed_UnknownFutureField_IsIgnored()
        {
            // Forward-compat lock-in for the JavaScriptSerializer choice (plan 4.4):
            // a payload with an extra field must round-trip without throwing.
            File.WriteAllText(_path,
                "{\"schemaVersion\":1,\"pinned\":[\"cs\"],\"enabled\":[\"cs\"],\"futureField\":\"hello\"}",
                new UTF8Encoding(false));

            var s = LanguageSettings.LoadOrSeed(_path, KnownTags, DefaultsPhase1);

            Assert.IsTrue(s.IsPinned("cs"));
            Assert.AreEqual(1, s.Pinned.Count);
        }

        [TestMethod]
        public void LoadOrSeed_DuplicateEntries_DeduplicatedFirstWins()
        {
            File.WriteAllText(_path,
                "{\"schemaVersion\":1,\"pinned\":[\"cs\",\"cs\",\"py\"],\"enabled\":[\"cs\",\"py\",\"py\"]}",
                new UTF8Encoding(false));

            var s = LanguageSettings.LoadOrSeed(_path, KnownTags, DefaultsPhase1);

            Assert.AreEqual(2, s.Pinned.Count);
            Assert.AreEqual("cs", s.Pinned[0]);
            Assert.AreEqual("py", s.Pinned[1]);
        }

        [TestMethod]
        public void LoadOrSeed_PinnedImpliesEnabled()
        {
            // The on-disk shape can technically omit a pinned tag from enabled;
            // the in-memory invariant must still hold.
            File.WriteAllText(_path,
                "{\"schemaVersion\":1,\"pinned\":[\"cs\"],\"enabled\":[]}",
                new UTF8Encoding(false));

            var s = LanguageSettings.LoadOrSeed(_path, KnownTags, DefaultsPhase1);

            Assert.IsTrue(s.IsPinned("cs"));
            Assert.IsTrue(s.IsEnabled("cs"));
        }

        // --- SaveAtomically ----------------------------------------------------

        [TestMethod]
        public void SaveAtomically_WhenDestinationDoesNotExist_UsesFileMove()
        {
            // Seed an in-memory instance via LoadOrSeed on a fresh dir.
            var s = LanguageSettings.LoadOrSeed(_path, KnownTags, DefaultsPhase1);

            File.Delete(_path);
            Assert.IsFalse(File.Exists(_path));

            s.SaveAtomically(_path);

            Assert.IsTrue(File.Exists(_path));
            // Tmp file must not linger.
            Assert.IsFalse(File.Exists(_path + ".tmp"));
        }

        [TestMethod]
        public void SaveAtomically_WhenDestinationExists_ReplacesIt()
        {
            // First save creates it.
            var s = LanguageSettings.LoadOrSeed(_path, KnownTags, DefaultsPhase1);
            s.SaveAtomically(_path);
            var firstBytes = File.ReadAllBytes(_path);

            // Mutate and save again - the file must be replaced, not appended.
            s.Pin("md", 0);
            s.SaveAtomically(_path);

            var secondBytes = File.ReadAllBytes(_path);
            CollectionAssert.AreNotEqual(firstBytes, secondBytes);
            Assert.IsFalse(File.Exists(_path + ".tmp"));
        }

        [TestMethod]
        public void SaveAtomically_CreatesParentDirectoryIfMissing()
        {
            var nested = Path.Combine(_tempDir, "nested", "deep", "languages.json");
            Assert.IsFalse(Directory.Exists(Path.GetDirectoryName(nested)));

            var s = LanguageSettings.LoadOrSeed(null, KnownTags, DefaultsPhase1); // in-memory only
            s.SaveAtomically(nested);

            Assert.IsTrue(File.Exists(nested));
        }

        [TestMethod]
        public void RoundTrip_PreservesPinnedAndEnabled()
        {
            var s1 = LanguageSettings.LoadOrSeed(_path, KnownTags, DefaultsPhase1);
            s1.Enable("go");
            s1.Pin("md", 0);
            s1.SaveAtomically(_path);

            var s2 = LanguageSettings.LoadOrSeed(_path, KnownTags, DefaultsPhase1);

            Assert.IsTrue(s2.IsPinned("md"));
            Assert.AreEqual("md", s2.Pinned[0]);
            Assert.IsTrue(s2.IsEnabled("go"));
            Assert.IsFalse(s2.IsPinned("go"));
            Assert.IsTrue(s2.IsVisibleInMoreMenu("go"));
        }

        // --- Mutation invariants ----------------------------------------------

        [TestMethod]
        public void Pin_UnknownPosition_ClampsToEnd()
        {
            var s = LanguageSettings.LoadOrSeed(null, null, DefaultsPhase1);
            s.Pin("brand-new-tag", 9999);

            Assert.AreEqual("brand-new-tag", s.Pinned[s.Pinned.Count - 1]);
            Assert.IsTrue(s.IsEnabled("brand-new-tag"));
        }

        [TestMethod]
        public void Unpin_KeepsEnabled()
        {
            var s = LanguageSettings.LoadOrSeed(null, null, DefaultsPhase1);
            s.Unpin("cs");

            Assert.IsFalse(s.IsPinned("cs"));
            Assert.IsTrue(s.IsEnabled("cs"), "Unpin must not disable.");
        }

        [TestMethod]
        public void Disable_RemovesFromBothLists()
        {
            var s = LanguageSettings.LoadOrSeed(null, null, DefaultsPhase1);
            s.Disable("cs");

            Assert.IsFalse(s.IsPinned("cs"));
            Assert.IsFalse(s.IsEnabled("cs"));
        }

        [TestMethod]
        public void Reorder_AppliesNewOrder_AndImpliesEnabled()
        {
            var s = LanguageSettings.LoadOrSeed(null, null, DefaultsPhase1);
            s.Reorder(new List<string> { "py", "cs" });

            Assert.AreEqual(2, s.Pinned.Count);
            Assert.AreEqual("py", s.Pinned[0]);
            Assert.AreEqual("cs", s.Pinned[1]);
            Assert.IsTrue(s.IsEnabled("py"));
            Assert.IsTrue(s.IsEnabled("cs"));
        }

        // --- CloneForEditing (regression: default-pinned tag could not be disabled) -----------

        [TestMethod]
        public void CloneForEditing_DisabledDefaultTag_StaysDisabled()
        {
            // The user disabled "css" (a Phase 1 default-pinned tag) in a prior session.
            // The live state therefore has "css" in neither Pinned nor Enabled.
            var live = LanguageSettings.LoadOrSeed(null, null, DefaultsPhase1);
            live.Disable("css");
            Assert.IsFalse(live.IsEnabled("css"), "Precondition: live state must not have css enabled.");

            // Reopening the dialog calls CloneForEditing with the global DefaultsPhase1 fallback.
            // Pre-fix, the clone was seeded from DefaultsPhase1 (which still contains "css")
            // and Reorder only cleared _pinned, leaving _enabled untouched. The clone would
            // therefore reappear with css enabled.
            var clone = NoteHighlightAddin.LanguageSettingsForm.CloneForEditing(live, DefaultsPhase1);

            Assert.IsFalse(clone.IsPinned("css"), "Clone must not re-pin css.");
            Assert.IsFalse(clone.IsEnabled("css"), "Clone must not resurrect a disabled default-pinned tag.");
        }

        [TestMethod]
        public void CloneForEditing_PreservesPinnedOrderAndExtraEnabled()
        {
            // Live state: pinned re-ordered, and an extra non-default tag enabled (not pinned).
            var live = LanguageSettings.LoadOrSeed(null, null, DefaultsPhase1);
            live.Reorder(new List<string> { "py", "cs", "js" });
            live.Enable("go");

            var clone = NoteHighlightAddin.LanguageSettingsForm.CloneForEditing(live, DefaultsPhase1);

            CollectionAssert.AreEqual(new[] { "py", "cs", "js" }, clone.Pinned.ToArray());
            Assert.IsTrue(clone.IsEnabled("go"), "Extra enabled tag must survive cloning.");
            Assert.IsFalse(clone.IsPinned("go"));
            // No leakage of un-pinned defaults into pinned.
            Assert.IsFalse(clone.IsPinned("css"));
            Assert.IsFalse(clone.IsPinned("html"));
        }

        [TestMethod]
        public void CloneForEditing_NullLive_FallsBackToDefaults()
        {
            var clone = NoteHighlightAddin.LanguageSettingsForm.CloneForEditing(null, DefaultsPhase1);

            Assert.AreEqual(DefaultsPhase1.Length, clone.Pinned.Count);
            foreach (var t in DefaultsPhase1)
            {
                Assert.IsTrue(clone.IsPinned(t));
                Assert.IsTrue(clone.IsEnabled(t));
            }
        }
    }
}
