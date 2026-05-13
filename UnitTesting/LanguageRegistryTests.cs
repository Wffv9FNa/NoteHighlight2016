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
        public void Initialise_NoArgOverload_UsesEmbeddedOrOnDisk()
        {
            // The no-arg overload resolves the addin directory via
            // AddIn.GetAddinDirectory() and then falls back to the embedded
            // copy. Under the test host that path may or may not contain a
            // ribbon.xml; either way the registry must parse the real
            // production languages.
            LanguageRegistry.Initialise();

            Assert.IsTrue(LanguageRegistry.IsInitialised);
            Assert.IsTrue(LanguageRegistry.All.Count >= 13,
                "Expected at least 13 languages from the no-arg Initialise path; got " + LanguageRegistry.All.Count);
            Assert.IsNotNull(LanguageRegistry.ByTag("cs"));
        }
    }
}
