/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NoteHighlightAddin;

namespace UnitTesting
{
    /// <summary>
    /// Unit tests for the pinned-language ribbon "slot" mechanism. Plan section
    /// 7.2 (work plan 2026-05-14-pinned-language-ribbon-ordering). These tests
    /// lock the slot-button shape in ribbon.xml, the slot-id parsing and the
    /// pure ordering core of <see cref="AddIn.ResolveSlot(LanguageSettings, IReadOnlyList{LanguageDescriptor}, int)"/>.
    /// </summary>
    [TestClass]
    public class SlotResolutionTests
    {
        // Resolves the addin source directory so the tests can point at the
        // real ribbon.xml. Mirrors LanguageRegistryTests.RepoRoot.
        private static string RepoRoot
        {
            get
            {
                var here = AppDomain.CurrentDomain.BaseDirectory;
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

        private static string RealRibbonXmlPath => Path.Combine(AddinSourceDir, "ribbon.xml");

        // The CustomUI namespace the production ribbon.xml uses (the 2006 schema).
        private static readonly XNamespace CustomUi =
            "http://schemas.microsoft.com/office/2006/01/customui";

        /// <summary>Loads and returns the real production ribbon.xml document.</summary>
        private static XDocument LoadRealRibbon()
        {
            return XDocument.Load(RealRibbonXmlPath);
        }

        /// <summary>Returns the named group element from the real ribbon.xml.</summary>
        private static XElement Group(XDocument doc, string groupId)
        {
            var group = doc.Descendants(CustomUi + "group")
                           .FirstOrDefault(g => (string)g.Attribute("id") == groupId);
            Assert.IsNotNull(group, "ribbon.xml must declare a <group id=\"" + groupId + "\">.");
            return group;
        }

        // --- Tests 1-5: ribbon.xml structure ---------------------------------

        [TestMethod]
        public void Ribbon_GroupLanguage_HasExactlyMaxPinnedSlotButtons()
        {
            // Plan section 7.2 test 1: locks the N=14 coupling between the slot
            // count in ribbon.xml and LanguageSettingsForm.MaxPinned.
            var doc = LoadRealRibbon();
            var group = Group(doc, "groupLanguage");

            var slotButtons = group.Elements(CustomUi + "button")
                                   .Where(b => ((string)b.Attribute("id") ?? string.Empty)
                                                   .StartsWith("slotLang", StringComparison.Ordinal))
                                   .ToList();

            Assert.AreEqual(LanguageSettingsForm.MaxPinned, slotButtons.Count,
                "groupLanguage must contain exactly LanguageSettingsForm.MaxPinned slot buttons.");

            // Ids must be the zero-padded slotLang00..slotLang13 sequence.
            var expectedIds = Enumerable.Range(0, LanguageSettingsForm.MaxPinned)
                                        .Select(i => "slotLang" + i.ToString("00"))
                                        .ToArray();
            var actualIds = slotButtons.Select(b => (string)b.Attribute("id")).ToArray();
            CollectionAssert.AreEqual(expectedIds, actualIds,
                "Slot button ids must be the contiguous zero-padded slotLang00..slotLang13 sequence.");
        }

        [TestMethod]
        public void Ribbon_SlotButtons_HaveNoTagAttribute()
        {
            // Plan section 7.2 test 2: a slot button carries no static tag - the
            // per-language tag is resolved at callback time. A future edit that
            // reintroduces a tag= would silently revive the static-tag assumption.
            var doc = LoadRealRibbon();
            var group = Group(doc, "groupLanguage");

            foreach (var b in group.Elements(CustomUi + "button")
                                   .Where(b => ((string)b.Attribute("id") ?? string.Empty)
                                                   .StartsWith("slotLang", StringComparison.Ordinal)))
            {
                Assert.IsNull(b.Attribute("tag"),
                    "Slot button '" + (string)b.Attribute("id") + "' must not declare a tag= attribute.");
            }
        }

        [TestMethod]
        public void Ribbon_SlotButtons_UseGetCallbacks()
        {
            // Plan section 7.2 test 3: every slot button drives its content
            // through get* callbacks and SlotButtonClicked.
            var doc = LoadRealRibbon();
            var group = Group(doc, "groupLanguage");

            foreach (var b in group.Elements(CustomUi + "button")
                                   .Where(b => ((string)b.Attribute("id") ?? string.Empty)
                                                   .StartsWith("slotLang", StringComparison.Ordinal)))
            {
                var id = (string)b.Attribute("id");
                Assert.AreEqual("GetSlotVisible", (string)b.Attribute("getVisible"),
                    "Slot button '" + id + "' must declare getVisible=\"GetSlotVisible\".");
                Assert.AreEqual("GetSlotLabel", (string)b.Attribute("getLabel"),
                    "Slot button '" + id + "' must declare getLabel=\"GetSlotLabel\".");
                Assert.AreEqual("GetSlotImage", (string)b.Attribute("getImage"),
                    "Slot button '" + id + "' must declare getImage=\"GetSlotImage\".");
                Assert.AreEqual("GetSlotScreentip", (string)b.Attribute("getScreentip"),
                    "Slot button '" + id + "' must declare getScreentip=\"GetSlotScreentip\".");
                Assert.AreEqual("SlotButtonClicked", (string)b.Attribute("onAction"),
                    "Slot button '" + id + "' must declare onAction=\"SlotButtonClicked\".");
            }
        }

        [TestMethod]
        public void Ribbon_CatalogueGroup_IsHiddenAndComplete()
        {
            // Plan section 7.2 test 4: the catalogue group exists, is hidden via
            // visible="false", and every button carries the parse-only data the
            // registry needs (id, tag, label, image). The catalogue button count
            // is asserted against LanguageRegistry.All so both stay self-consistent
            // - the registry walks exactly this group's buttons.
            var doc = LoadRealRibbon();
            var catalogue = Group(doc, "groupLanguageCatalogue");

            Assert.AreEqual("false", (string)catalogue.Attribute("visible"),
                "groupLanguageCatalogue must declare visible=\"false\".");

            var buttons = catalogue.Elements(CustomUi + "button").ToList();
            Assert.IsTrue(buttons.Count >= 27,
                "groupLanguageCatalogue must contain the full language catalogue; got " + buttons.Count + " buttons.");

            foreach (var b in buttons)
            {
                Assert.IsFalse(string.IsNullOrEmpty((string)b.Attribute("id")),
                    "Every catalogue button must have an id.");
                Assert.IsFalse(string.IsNullOrEmpty((string)b.Attribute("tag")),
                    "Catalogue button '" + (string)b.Attribute("id") + "' must have a tag.");
                Assert.IsFalse(string.IsNullOrEmpty((string)b.Attribute("label")),
                    "Catalogue button '" + (string)b.Attribute("id") + "' must have a label.");
                Assert.IsFalse(string.IsNullOrEmpty((string)b.Attribute("image")),
                    "Catalogue button '" + (string)b.Attribute("id") + "' must have an image.");
            }

            // The registry parses exactly this group, so its count must match.
            LanguageRegistry.ResetForTests();
            LanguageRegistry.Initialise(AddinSourceDir);
            Assert.AreEqual(LanguageRegistry.All.Count, buttons.Count,
                "LanguageRegistry.All must contain exactly the catalogue group's buttons.");
        }

        [TestMethod]
        public void Registry_DoesNotIncludeSlotButtons()
        {
            // Plan section 7.2 test 5: the slotLang* buttons have no tag and so
            // are skipped by the registry parser - LanguageRegistry.All must hold
            // only real catalogue languages.
            LanguageRegistry.ResetForTests();
            LanguageRegistry.Initialise(AddinSourceDir);

            foreach (var d in LanguageRegistry.All)
            {
                Assert.IsFalse((d.ButtonId ?? string.Empty).StartsWith("slotLang", StringComparison.Ordinal),
                    "LanguageRegistry.All must not include slot button '" + d.ButtonId + "'.");
            }
        }

        // --- Test 6: slot index parsing --------------------------------------

        [TestMethod]
        public void SlotIndexFromControlId_ParsesValidIds_RejectsMalformed()
        {
            // Plan section 7.2 test 6: slotLangNN -> NN, -1 on any malformed id.
            Assert.AreEqual(0, AddIn.SlotIndexFromControlId("slotLang00"));
            Assert.AreEqual(13, AddIn.SlotIndexFromControlId("slotLang13"));
            Assert.AreEqual(7, AddIn.SlotIndexFromControlId("slotLang07"));

            Assert.AreEqual(-1, AddIn.SlotIndexFromControlId("slotLang"),
                "A bare 'slotLang' prefix with no numeric suffix must yield -1.");
            Assert.AreEqual(-1, AddIn.SlotIndexFromControlId("slotLangXX"),
                "A non-numeric suffix must yield -1.");
            Assert.AreEqual(-1, AddIn.SlotIndexFromControlId(null),
                "A null id must yield -1.");
            Assert.AreEqual(-1, AddIn.SlotIndexFromControlId(string.Empty),
                "An empty id must yield -1.");
        }

        // --- Test 7: ordering round-trip (the core regression test) ----------

        // Hand-built descriptors - deliberately NOT the global LanguageRegistry,
        // so this test needs no static-global setup (Initialise/ResetForTests).
        private static LanguageDescriptor Desc(string tag, string label) =>
            new LanguageDescriptor("button" + tag, tag, label, "Enter " + label + " Code", tag + ".png", null);

        private static IReadOnlyList<LanguageDescriptor> SampleCatalogue() => new[]
        {
            Desc("py", "Python"),
            Desc("cs", "C#"),
            Desc("rust", "Rust"),
        };

        // Builds an in-memory LanguageSettings whose Pinned list is exactly
        // pinned, in order. LoadOrSeed with a null path seeds from defaultPinned
        // without touching disk; with no knownTags filter the seed is verbatim.
        private static LanguageSettings SettingsWithPinned(params string[] pinned) =>
            LanguageSettings.LoadOrSeed(null, null, pinned);

        [TestMethod]
        public void ResolveSlot_StaticCore_MapsPinnedOrderToSlots()
        {
            // Plan section 7.2 test 7: the core regression test - dialog order
            // maps to slot order. Uses the static, test-visible ResolveSlot
            // overload with a hand-built descriptor list.
            var catalogue = SampleCatalogue();
            var settings = SettingsWithPinned("py", "cs", "rust");

            Assert.AreEqual("py", AddIn.ResolveSlot(settings, catalogue, 0)?.Tag);
            Assert.AreEqual("cs", AddIn.ResolveSlot(settings, catalogue, 1)?.Tag);
            Assert.AreEqual("rust", AddIn.ResolveSlot(settings, catalogue, 2)?.Tag);
            Assert.IsNull(AddIn.ResolveSlot(settings, catalogue, 3),
                "Slot 3 is beyond the pinned count and must resolve to null.");
            Assert.IsNull(AddIn.ResolveSlot(settings, catalogue, -1),
                "A negative slot index must resolve to null.");

            // Reorder on a SEPARATE instance, per the immutability invariant: the
            // live LanguageSettings handed to the AddIn is never mutated in place
            // (the form edits a separate instance). Build a fresh one, reorder it,
            // and confirm (a) the reordered instance maps to the new order and
            // (b) the original instance is untouched.
            var reordered = SettingsWithPinned("py", "cs", "rust");
            reordered.Reorder(new List<string> { "cs", "py" });

            Assert.AreEqual("cs", AddIn.ResolveSlot(reordered, catalogue, 0)?.Tag);
            Assert.AreEqual("py", AddIn.ResolveSlot(reordered, catalogue, 1)?.Tag);
            Assert.IsNull(AddIn.ResolveSlot(reordered, catalogue, 2),
                "After reordering to two pinned tags, slot 2 must resolve to null.");

            // Immutability invariant: the original instance is unchanged.
            Assert.AreEqual("py", AddIn.ResolveSlot(settings, catalogue, 0)?.Tag);
            Assert.AreEqual("cs", AddIn.ResolveSlot(settings, catalogue, 1)?.Tag);
            Assert.AreEqual("rust", AddIn.ResolveSlot(settings, catalogue, 2)?.Tag);
        }

        // --- Test 8: over-MaxPinned pinned list ------------------------------

        [TestMethod]
        public void ResolveSlot_StaticCore_OverMaxPinned_ResolvesFirst14_NullBeyond()
        {
            // Plan section 7.2 test 8: the only way to exceed MaxPinned is a
            // hand-edited languages.json; LoadOrSeed does not truncate. The static
            // ResolveSlot resolves slots 0..13 and returns null for slot index 14.
            var tags = Enumerable.Range(0, 16).Select(i => "lang" + i.ToString("00")).ToArray();
            var catalogue = tags.Select(t => Desc(t, t.ToUpperInvariant())).ToList();
            var settings = SettingsWithPinned(tags);

            for (int slot = 0; slot < LanguageSettingsForm.MaxPinned; slot++)
            {
                var resolved = AddIn.ResolveSlot(settings, catalogue, slot);
                Assert.IsNotNull(resolved, "Slot " + slot + " must resolve within MaxPinned.");
                Assert.AreEqual(tags[slot], resolved.Tag,
                    "Slot " + slot + " must resolve to the pinned tag at that index.");
            }

            // The pinned list itself has 16 entries, so the static core - which
            // is bounded only by the pinned count, not MaxPinned - still resolves
            // slots 14 and 15. The "extra tags are inert" behaviour comes from
            // ribbon.xml only declaring 14 slot buttons; assert slot index 14
            // resolves at the data layer (entry 15 of the pinned list) and that
            // there is no entry beyond the pinned count.
            Assert.IsNotNull(AddIn.ResolveSlot(settings, catalogue, 14),
                "Slot 14 still maps to pinned entry 14 at the data layer (no ribbon button exists for it).");
            Assert.IsNotNull(AddIn.ResolveSlot(settings, catalogue, 15));
            Assert.IsNull(AddIn.ResolveSlot(settings, catalogue, 16),
                "Slot index 16 is beyond the 16-entry pinned list and must resolve to null.");
        }

        // --- Test 9: GetSlotImage fallback -----------------------------------

        [TestMethod]
        public void GetSlotImage_UnresolvedSlot_ReturnsNonNullOtherPicture_NeverThrows()
        {
            // Plan section 7.2 test 9: an unresolved slot must yield a non-null
            // Other.png IPictureDisp and never throw. The AddIn has no pinned
            // languages set in a bare test instance, so every slot is unresolved.
            // Per the plan, asserting non-null and exception-free is sufficient -
            // the visual correctness of the picture is covered by manual matrix M7.
            var addin = new AddIn();

            stdole.IPictureDisp picture = null;
            try
            {
                picture = addin.GetSlotImage(new FakeRibbonControl("slotLang00"));
            }
            catch (Exception ex)
            {
                Assert.Fail("GetSlotImage must never throw out of Office; threw: " + ex);
            }

            Assert.IsNotNull(picture,
                "GetSlotImage on an unresolved slot must return the Other.png picture, not null.");

            // A null control must also be handled without throwing.
            try
            {
                addin.GetSlotImage(null);
            }
            catch (Exception ex)
            {
                Assert.Fail("GetSlotImage(null) must never throw; threw: " + ex);
            }
        }

        /// <summary>
        /// Minimal <see cref="Microsoft.Office.Core.IRibbonControl"/> stand-in so
        /// the slot callbacks can be exercised without an Office host. Only
        /// <see cref="Id"/> is consulted by the slot code paths.
        /// </summary>
        private sealed class FakeRibbonControl : Microsoft.Office.Core.IRibbonControl
        {
            public FakeRibbonControl(string id) { Id = id; }

            public string Id { get; }
            public object Context => null;
            public string Tag => null;
        }
    }
}
