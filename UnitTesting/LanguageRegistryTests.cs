/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NoteHighlightAddin;

namespace UnitTesting
{
    /// <summary>
    /// Unit tests for <see cref="LanguageRegistry"/>. Plan section 9.1.
    /// The registry uses static state, so each test that exercises Initialise
    /// resets it first via <see cref="LanguageRegistry.ResetForTests"/>.
    /// </summary>
    [TestClass]
    public class LanguageRegistryTests
    {
        // Resolves the addin source directory so the tests can point at the
        // real ribbon.xml. The test DLL ships under
        // UnitTesting\bin\<Configuration>\ but ribbon.xml lives at
        // NoteHighlightAddin\ribbon.xml relative to the repo root.
        private static string RepoRoot
        {
            get
            {
                var here = AppDomain.CurrentDomain.BaseDirectory;
                // Walk up until we find NoteHighlightAddin\ribbon.xml.
                var dir = new DirectoryInfo(here);
                while (dir != null)
                {
                    var candidate = Path.Combine(dir.FullName, "NoteHighlightAddin", "ribbon.xml");
                    if (File.Exists(candidate)) return dir.FullName;
                    dir = dir.Parent;
                }
                throw new InvalidOperationException("Could not locate repo root containing NoteHighlightAddin\\ribbon.xml from " + here);
            }
        }

        private static string AddinSourceDir => Path.Combine(RepoRoot, "NoteHighlightAddin");

        [TestInitialize]
        public void ResetRegistry() => LanguageRegistry.ResetForTests();

        [TestMethod]
        public void Initialise_RealRibbonXml_ParsesAllLanguageButtons()
        {
            LanguageRegistry.Initialise(AddinSourceDir);

            Assert.IsTrue(LanguageRegistry.IsInitialised, "Registry should report initialised after Initialise.");
            Assert.IsTrue(LanguageRegistry.All.Count >= 13,
                "Expected at least the 13 visible-true buttons to be parsed; got " + LanguageRegistry.All.Count);

            // Every Phase 1 default-pinned tag must resolve via ribbon.xml.
            foreach (var tag in LanguageRegistry.DefaultPinned)
            {
                Assert.IsNotNull(LanguageRegistry.ByTag(tag), "ribbon.xml is missing default-pinned tag '" + tag + "'.");
            }
        }

        [TestMethod]
        public void DefaultPinned_PreservesPhase1VisibleTrueSet()
        {
            // Plan section 11 - Phase 1 seed must equal today's visible="true" list,
            // not the section 3.2 refresh (cs/sql/py/js/...). Locking this here makes
            // sure a later refactor does not silently regress the appearance contract.
            var expected = new[]
            {
                "cs", "sql", "css", "js", "html", "xml", "java",
                "php", "perl", "py", "ruby", "c", "ps1",
            };

            CollectionAssert.AreEqual(expected, LanguageRegistry.DefaultPinned.ToArray());
        }

        [TestMethod]
        public void ByTag_KnownTag_ReturnsDescriptor()
        {
            LanguageRegistry.Initialise(AddinSourceDir);

            var cs = LanguageRegistry.ByTag("cs");
            Assert.IsNotNull(cs, "ribbon.xml must declare a 'cs' tag.");
            Assert.AreEqual("buttonCSharp", cs.ButtonId);
            Assert.AreEqual("C#", cs.Label);
        }

        [TestMethod]
        public void ByTag_UnknownTag_ReturnsNull()
        {
            LanguageRegistry.Initialise(AddinSourceDir);
            Assert.IsNull(LanguageRegistry.ByTag("definitely-not-a-real-tag-zzz"));
        }

        [TestMethod]
        public void Initialise_MissingRibbonXml_DoesNotThrowAndFallsBackToEmbedded()
        {
            // Bug 13.1 fix changed the contract here: when no on-disk ribbon.xml
            // is found at addinDirectory, Initialise now falls back to the
            // embedded Resources.ribbon copy rather than leaving All empty.
            // (The "leave empty" behaviour only applies when BOTH sources fail,
            // which is exercised by an out-of-process unit test with a mocked
            // resource ManagerStream - out of scope here.)
            var bogus = Path.Combine(Path.GetTempPath(), "nh-tests-missing-" + Guid.NewGuid().ToString("N"));

            // Must not throw.
            LanguageRegistry.Initialise(bogus);

            Assert.IsTrue(LanguageRegistry.IsInitialised);
            Assert.IsTrue(LanguageRegistry.All.Count >= 13,
                "Embedded fallback should populate the registry; got " + LanguageRegistry.All.Count);
        }

        [TestMethod]
        public void Initialise_NullDirectory_DoesNotThrow()
        {
            // Same contract as the test above: null directory means "skip the
            // on-disk hot-edit step, go straight to the embedded copy".
            LanguageRegistry.Initialise(null);

            Assert.IsTrue(LanguageRegistry.IsInitialised);
            Assert.IsTrue(LanguageRegistry.All.Count >= 13,
                "Embedded fallback should populate the registry even with a null addinDirectory; got " + LanguageRegistry.All.Count);
        }

        [TestMethod]
        public void Initialise_IsIdempotent()
        {
            LanguageRegistry.Initialise(AddinSourceDir);
            var firstCount = LanguageRegistry.All.Count;

            // Pointing at a missing path on the second call must NOT clobber the first init.
            LanguageRegistry.Initialise(Path.Combine(Path.GetTempPath(), "nh-tests-second-" + Guid.NewGuid().ToString("N")));

            Assert.AreEqual(firstCount, LanguageRegistry.All.Count, "Second Initialise should be a no-op.");
        }

        [TestMethod]
        public void Initialise_EmbeddedFallback_PopulatesRealLanguages()
        {
            // Bug 13.1 / Option B: when no on-disk ribbon.xml is present next to
            // the addin directory we are told to consult, Initialise must fall
            // back to the embedded Resources.ribbon copy and still expose the
            // canonical language set. Pointing at a guaranteed-missing temp
            // directory exercises that fallback path.
            var bogus = Path.Combine(Path.GetTempPath(), "nh-tests-embedded-" + Guid.NewGuid().ToString("N"));
            Assert.IsFalse(File.Exists(Path.Combine(bogus, "ribbon.xml")), "Sanity: bogus dir must not contain ribbon.xml.");

            LanguageRegistry.Initialise(bogus);

            Assert.IsTrue(LanguageRegistry.IsInitialised, "Registry should report initialised after embedded fallback.");
            Assert.IsTrue(LanguageRegistry.All.Count >= 13,
                "Embedded ribbon resource should yield at least the 13 production languages; got " + LanguageRegistry.All.Count);

            // Spot-check a couple of canonical tags that ship in the embedded ribbon.xml.
            Assert.IsNotNull(LanguageRegistry.ByTag("cs"), "Embedded ribbon.xml must declare 'cs'.");
            Assert.IsNotNull(LanguageRegistry.ByTag("py"), "Embedded ribbon.xml must declare 'py'.");
        }

        [TestMethod]
        public void Initialise_NewLanguages_AreAllRegistered()
        {
            // Plan section 5.1: locks the contract that all six newly-added
            // ribbon buttons resolve through LanguageRegistry. A failure here
            // means ribbon.xml was edited without updating the registry parser,
            // or the new <button> rows are missing required attributes.
            LanguageRegistry.Initialise(AddinSourceDir);

            var expected = new[] { "yaml", "rust", "kotlin", "dockerfile", "tex", "autohotkey" };
            foreach (var tag in expected)
            {
                var d = LanguageRegistry.ByTag(tag);
                Assert.IsNotNull(d, "ribbon.xml must declare a button with tag='" + tag + "'.");
                Assert.IsFalse(string.IsNullOrEmpty(d.Label), "tag '" + tag + "' must have a label.");
                Assert.IsFalse(string.IsNullOrEmpty(d.Image), "tag '" + tag + "' must have an image.");
            }
        }

        [TestMethod]
        public void RemovedLanguages_AreNotRegistered()
        {
            // Plan section 4 Phase 1: locks the prune. A failure here means a
            // re-added button slipped back into ribbon.xml without a matching plan
            // entry, or the prune commit was partially reverted.
            LanguageRegistry.Initialise(AddinSourceDir);

            var removed = new[]
            {
                "clojure", "conf", "fsharp", "go", "haskell", "lisp",
                "logtalk", "make", "matlab", "pas", "r", "swift",
            };

            foreach (var tag in removed)
            {
                Assert.IsNull(LanguageRegistry.ByTag(tag),
                    "ribbon.xml must NOT declare a button with tag='" + tag + "' after the 4.0 prune.");
            }
        }

        [TestMethod]
        public void NewLanguages_HaveMatchingLangDefFiles()
        {
            // Plan section 5.2: build-time sanity check. Each new ribbon tag
            // must correspond to a *.lang file shipped under highlight\langDefs,
            // otherwise highlight.exe will reject --syntax=<tag> at runtime.
            var langDefsDir = Path.Combine(RepoRoot, "NoteHighlightAddin", "highlight", "langDefs");
            Assert.IsTrue(Directory.Exists(langDefsDir), "langDefs directory must exist.");

            var expected = new[] { "yaml", "rust", "kotlin", "dockerfile", "tex", "autohotkey" };
            foreach (var tag in expected)
            {
                var langFile = Path.Combine(langDefsDir, tag + ".lang");
                Assert.IsTrue(File.Exists(langFile), "Missing langDef file: " + langFile);
            }
        }

        [TestMethod]
        public void RegistryTags_AreAllUnique()
        {
            // Plan section 5.3: cheap insurance against a copy-paste typo
            // introducing a duplicate tag="..." in ribbon.xml. LanguageRegistry
            // groups its internal lookup by id, which would silently swallow a
            // second button with the same tag - this assertion catches it.
            LanguageRegistry.Initialise(AddinSourceDir);

            var tags = LanguageRegistry.All.Select(d => d.Tag).ToList();
            var distinct = tags.Distinct(StringComparer.Ordinal).Count();
            Assert.AreEqual(tags.Count, distinct,
                "Duplicate tag in ribbon.xml. Tags: " + string.Join(", ", tags));
        }

        [TestMethod]
        public void Initialise_NoArgOverload_UsesEmbeddedOrOnDisk()
        {
            // The no-arg overload resolves the addin directory via
            // AddIn.GetAddinDirectory() and then falls back to the embedded
            // copy. Under the test host that path may or may not contain a
            // ribbon.override.xml; either way the registry must parse the
            // real production languages.
            LanguageRegistry.Initialise();

            Assert.IsTrue(LanguageRegistry.IsInitialised);
            Assert.IsTrue(LanguageRegistry.All.Count >= 13,
                "Expected at least 13 languages from the no-arg Initialise path; got " + LanguageRegistry.All.Count);
            Assert.IsNotNull(LanguageRegistry.ByTag("cs"));
        }

        // --- Bug 13.1 followup (2026-05-13 evening) ----------------------
        // The following tests lock the contract for the "stale on-disk
        // ribbon.xml beats the embedded copy" fix. They use isolated temp
        // directories so the real addin source tree is not touched.

        /// <summary>
        /// Minimal valid ribbon.xml fragment that LanguageRegistry can parse
        /// to produce a known descriptor set. Used by the fixture-based tests
        /// below so each one is independent of ribbon.xml drift.
        /// </summary>
        private static string FixtureRibbon(params (string id, string tag, string label)[] buttons)
        {
            var sb = new System.Text.StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            sb.Append("<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">");
            sb.Append("<ribbon>");
            sb.Append("<tabs>");
            sb.Append("<tab id=\"tabFixture\" label=\"Fixture\">");
            sb.Append("<group id=\"groupLanguage\" label=\"Language\">");
            foreach (var b in buttons)
            {
                sb.Append("<button id=\"").Append(b.id).Append("\"")
                  .Append(" tag=\"").Append(b.tag).Append("\"")
                  .Append(" label=\"").Append(b.label).Append("\"")
                  .Append(" image=\"Other.png\"")
                  .Append(" onAction=\"AddInButtonClicked\" />");
            }
            sb.Append("</group>");
            sb.Append("</tab>");
            sb.Append("</tabs>");
            sb.Append("</ribbon>");
            sb.Append("</customUI>");
            return sb.ToString();
        }

        private static string CreateTempDir(string prefix)
        {
            var dir = Path.Combine(Path.GetTempPath(), prefix + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            return dir;
        }

        [TestMethod]
        public void Initialise_LegacyRibbonXmlPresent_IsIgnoredAndEmbeddedWins()
        {
            // Bug 13.1 followup safety net: a legacy ribbon.xml left behind by
            // an old MSI (Permanent=TRUE component, never uninstalled) MUST
            // NOT be read - it would beat the embedded canonical copy with
            // stale content. The registry falls back to the embedded copy and
            // emits a Trace warning naming the resolved legacy path.
            var dir = CreateTempDir("nh-tests-legacy");
            try
            {
                // Write a deliberately tiny legacy ribbon.xml with one button
                // only - far smaller than the real embedded copy, so we can
                // tell the embedded copy won by counting.
                var legacy = Path.Combine(dir, LanguageRegistry.LegacyOnDiskRibbonFileName);
                File.WriteAllText(legacy, FixtureRibbon(("buttonStale", "stale-only", "Stale")));

                using (var listener = new TraceCaptureListener())
                {
                    LanguageRegistry.Initialise(dir);

                    Assert.IsTrue(LanguageRegistry.IsInitialised);
                    Assert.IsTrue(LanguageRegistry.All.Count >= 13,
                        "Embedded copy should have won; instead got " + LanguageRegistry.All.Count + " (would mean the legacy ribbon.xml was read).");
                    Assert.IsNull(LanguageRegistry.ByTag("stale-only"),
                        "The legacy ribbon.xml tag must NOT appear - the embedded copy should be canonical.");

                    // The Trace warning must name the resolved path so a user
                    // can grep their trace log for the bad file.
                    Assert.IsTrue(
                        listener.Messages.Any(m => m.IndexOf("legacy", StringComparison.OrdinalIgnoreCase) >= 0
                                                && m.IndexOf(legacy, StringComparison.OrdinalIgnoreCase) >= 0),
                        "Expected a Trace warning mentioning 'legacy' and the resolved path. Captured: " + string.Join(" | ", listener.Messages));
                }
            }
            finally
            {
                SafeDeleteDir(dir);
            }
        }

        [TestMethod]
        public void Initialise_OverrideXmlPresent_IsPreferredOverEmbedded()
        {
            // Hot-edit path: ribbon.override.xml next to the DLL wins over
            // the embedded copy. End-user installs never ship this filename,
            // so this only fires for devs.
            var dir = CreateTempDir("nh-tests-override");
            try
            {
                var overridePath = Path.Combine(dir, LanguageRegistry.OnDiskRibbonFileName);
                File.WriteAllText(overridePath, FixtureRibbon(
                    ("buttonAlpha", "alpha", "Alpha"),
                    ("buttonBeta", "beta", "Beta")));

                LanguageRegistry.Initialise(dir);

                Assert.IsTrue(LanguageRegistry.IsInitialised);
                Assert.AreEqual(2, LanguageRegistry.All.Count,
                    "Override file should be the only source; embedded copy must NOT have been used.");
                Assert.IsNotNull(LanguageRegistry.ByTag("alpha"));
                Assert.IsNotNull(LanguageRegistry.ByTag("beta"));
                Assert.IsNull(LanguageRegistry.ByTag("cs"),
                    "Embedded copy must not contribute when override is present.");
            }
            finally
            {
                SafeDeleteDir(dir);
            }
        }

        [TestMethod]
        public void Initialise_BothLegacyAndOverridePresent_OverrideWins_LegacyWarned()
        {
            // The worst case: a dev with a legacy ribbon.xml lying around AND
            // a deliberate ribbon.override.xml. The override must still win,
            // and the legacy file must still produce a Trace warning so the
            // dev knows to clean it up.
            var dir = CreateTempDir("nh-tests-both");
            try
            {
                var legacy = Path.Combine(dir, LanguageRegistry.LegacyOnDiskRibbonFileName);
                File.WriteAllText(legacy, FixtureRibbon(("buttonStale", "stale-only", "Stale")));

                var overridePath = Path.Combine(dir, LanguageRegistry.OnDiskRibbonFileName);
                File.WriteAllText(overridePath, FixtureRibbon(("buttonAlpha", "alpha", "Alpha")));

                using (var listener = new TraceCaptureListener())
                {
                    LanguageRegistry.Initialise(dir);

                    Assert.IsTrue(LanguageRegistry.IsInitialised);
                    Assert.AreEqual(1, LanguageRegistry.All.Count,
                        "Override file should win; got " + LanguageRegistry.All.Count + " buttons.");
                    Assert.IsNotNull(LanguageRegistry.ByTag("alpha"));
                    Assert.IsNull(LanguageRegistry.ByTag("stale-only"));

                    Assert.IsTrue(
                        listener.Messages.Any(m => m.IndexOf("legacy", StringComparison.OrdinalIgnoreCase) >= 0
                                                && m.IndexOf(legacy, StringComparison.OrdinalIgnoreCase) >= 0),
                        "Legacy-file warning still expected when an override file is also present. Captured: " + string.Join(" | ", listener.Messages));
                }
            }
            finally
            {
                SafeDeleteDir(dir);
            }
        }

        [TestMethod]
        public void OnDiskRibbonFileName_IsRibbonOverrideXml()
        {
            // Locks the filename contract: end-user MSIs must not ship this
            // name. If anyone renames it back to "ribbon.xml" this test will
            // shout (and so will the MSI safety: there is no MSI row matching
            // ribbon.override.xml, see Setup.vdproj).
            Assert.AreEqual("ribbon.override.xml", LanguageRegistry.OnDiskRibbonFileName);
            Assert.AreEqual("ribbon.xml", LanguageRegistry.LegacyOnDiskRibbonFileName);
        }

        private static void SafeDeleteDir(string dir)
        {
            try { if (Directory.Exists(dir)) Directory.Delete(dir, true); } catch { /* best-effort cleanup */ }
        }

        /// <summary>
        /// Captures Trace.TraceWarning / TraceError / TraceInformation output
        /// for the duration of a test. Registered globally via
        /// <see cref="System.Diagnostics.Trace.Listeners"/>; <c>Dispose</c>
        /// unregisters so the next test starts clean.
        /// </summary>
        private sealed class TraceCaptureListener : System.Diagnostics.TraceListener
        {
            public System.Collections.Generic.List<string> Messages { get; } = new System.Collections.Generic.List<string>();

            public TraceCaptureListener()
            {
                System.Diagnostics.Trace.Listeners.Add(this);
            }

            public override void Write(string message) { if (!string.IsNullOrEmpty(message)) Messages.Add(message); }
            public override void WriteLine(string message) { if (!string.IsNullOrEmpty(message)) Messages.Add(message); }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    System.Diagnostics.Trace.Listeners.Remove(this);
                }
                base.Dispose(disposing);
            }
        }
    }
}
