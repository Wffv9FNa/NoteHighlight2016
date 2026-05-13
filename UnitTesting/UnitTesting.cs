using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NoteHighlightAddin;
using GenerateHighlightContent;
using System.Xml.Linq;
using System.Configuration;
using System.Linq;
using System.Threading;

namespace UnitTesting
{
    [TestClass]
    public class UnitTesting
    {
        [TestMethod]
        public void FormatNewCode()
        {
            string htmlCode = Resource1.HTMLContent1;
            string[] pos = new string[] { "198.0", "950.3999633789062" };
            HighLightParameter param = new HighLightParameter();
            param.ShowLineNumber = true;
            param.HighlightColor = System.Drawing.Color.FromArgb(240, 240, 240);
            XElement outline = null;


            AddIn addIn = new AddIn();
            addIn.ns = @"http://schemas.microsoft.com/office/onenote/2013/onenote";


            //Arrange
            ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap { ExeConfigFilename = "Test.config" };
            Configuration configuration = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);

            HighLightSection config = configuration.GetSection("HighLightSection") as HighLightSection;


            XDocument output = addIn.InsertHighLightCode(htmlCode, pos, param, outline, config, false, false);

            Assert.AreEqual(Resource1.Output1, output.ToString(), false);
        }

        [TestMethod]
        public void FormatSelectedCode_AllSelected()
        {
            string htmlCode = Resource1.HTMLContent2;
            string[] pos = null;
            HighLightParameter param = new HighLightParameter();
            param.ShowLineNumber = true;
            param.HighlightColor = System.Drawing.Color.FromArgb(240, 240, 240);


            AddIn addIn = new AddIn();
            addIn.ns = @"http://schemas.microsoft.com/office/onenote/2013/onenote";

            var outline = XDocument.Parse(Resource1.Page2).Descendants(addIn.ns + "Outline")
                                   .Where(n => n.Attribute("selected") != null && (n.Attribute("selected").Value == "all" || n.Attribute("selected").Value == "partial"))
                                   .FirstOrDefault();

            //Arrange
            ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap { ExeConfigFilename = "Test.config" };
            Configuration configuration = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);

            HighLightSection config = configuration.GetSection("HighLightSection") as HighLightSection;

            bool selectedTextFormated = false;
            var page2Root = XDocument.Parse(Resource1.Page2).Root;
            addIn.GetSelectedText(page2Root, out selectedTextFormated);


                XDocument output = addIn.InsertHighLightCode(htmlCode, pos, param, outline, config, selectedTextFormated, addIn.IsSelectedTextInline(page2Root));

            Assert.AreEqual(Resource1.Output2, output.ToString(), false);
        }

        [TestMethod]
        public void FormatSelectedCode_WordSelected()
        {
            string htmlCode = Resource1.HTMLContent5;
            string[] pos = null;
            HighLightParameter param = new HighLightParameter();
            param.ShowLineNumber = true;
            param.HighlightColor = System.Drawing.Color.FromArgb(240, 240, 240);


            AddIn addIn = new AddIn();
            addIn.ns = @"http://schemas.microsoft.com/office/onenote/2013/onenote";

            var outline = XDocument.Parse(Resource1.Page5).Descendants(addIn.ns + "Outline")
                                   .Where(n => n.Attribute("selected") != null && (n.Attribute("selected").Value == "all" || n.Attribute("selected").Value == "partial"))
                                   .FirstOrDefault();

            //Arrange
            ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap { ExeConfigFilename = "Test.config" };
            Configuration configuration = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);

            HighLightSection config = configuration.GetSection("HighLightSection") as HighLightSection;

            bool selectedTextFormated = false;
            var page5Root = XDocument.Parse(Resource1.Page5).Root;
            addIn.GetSelectedText(page5Root, out selectedTextFormated);


            XDocument output = addIn.InsertHighLightCode(htmlCode, pos, param, outline, config, selectedTextFormated, addIn.IsSelectedTextInline(page5Root));

            Assert.AreEqual(Resource1.Output5, output.ToString(), false);
        }

        [TestMethod]
        public void FormatSelectedCode_LineSelected()
        {
            string htmlCode = Resource1.HTMLContent3;
            string[] pos = null;
            HighLightParameter param = new HighLightParameter();
            param.ShowLineNumber = true;
            param.HighlightColor = System.Drawing.Color.FromArgb(240, 240, 240);


            AddIn addIn = new AddIn();
            addIn.ns = @"http://schemas.microsoft.com/office/onenote/2013/onenote";

            var outline = XDocument.Parse(Resource1.Page3).Descendants(addIn.ns + "Outline")
                                   .Where(n => n.Attribute("selected") != null && (n.Attribute("selected").Value == "all" || n.Attribute("selected").Value == "partial"))
                                   .FirstOrDefault();

            //Arrange
            ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap { ExeConfigFilename = "Test.config" };
            Configuration configuration = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);

            HighLightSection config = configuration.GetSection("HighLightSection") as HighLightSection;

            bool selectedTextFormated = false;
            var page3Root = XDocument.Parse(Resource1.Page3).Root;
            addIn.GetSelectedText(page3Root, out selectedTextFormated);


            XDocument output = addIn.InsertHighLightCode(htmlCode, pos, param, outline, config, selectedTextFormated, addIn.IsSelectedTextInline(page3Root));

            Assert.AreEqual(Resource1.Output3, output.ToString(), false);
        }

        [TestMethod]
        public void FormatSelectedCode_PartialLineSelected()
        {
            string htmlCode = Resource1.HTMLContent6;
            string[] pos = null;
            HighLightParameter param = new HighLightParameter();
            param.ShowLineNumber = true;
            param.HighlightColor = System.Drawing.Color.FromArgb(240, 240, 240);


            AddIn addIn = new AddIn();
            addIn.ns = @"http://schemas.microsoft.com/office/onenote/2013/onenote";

            var outline = XDocument.Parse(Resource1.Page6).Descendants(addIn.ns + "Outline")
                                   .Where(n => n.Attribute("selected") != null && (n.Attribute("selected").Value == "all" || n.Attribute("selected").Value == "partial"))
                                   .FirstOrDefault();

            //Arrange
            ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap { ExeConfigFilename = "Test.config" };
            Configuration configuration = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);

            HighLightSection config = configuration.GetSection("HighLightSection") as HighLightSection;

            bool selectedTextFormated = false;
            var page6Root = XDocument.Parse(Resource1.Page6).Root;
            addIn.GetSelectedText(page6Root, out selectedTextFormated);


            XDocument output = addIn.InsertHighLightCode(htmlCode, pos, param, outline, config, selectedTextFormated, addIn.IsSelectedTextInline(page6Root));

            Assert.AreEqual(Resource1.Output6, output.ToString(), false);
        }

        [TestMethod]
        public void OneNoteMessageFilter_RetryRejectedCall_BackoffSequence()
        {
            // SERVERCALL_RETRYLATER is 2. The filter should return the configured backoff
            // sequence (100, 200, 400, ...) capped per-retry at 2000 ms and overall at
            // MaxRetryMilliseconds.
            const int SERVERCALL_RETRYLATER = 2;

            var filter = new OneNoteMessageFilter
            {
                BackoffStepMilliseconds = 100,
                MaxRetryMilliseconds = 30000,
            };

            int first = filter.RetryRejectedCall(IntPtr.Zero, 0, SERVERCALL_RETRYLATER);
            int second = filter.RetryRejectedCall(IntPtr.Zero, 0, SERVERCALL_RETRYLATER);
            int third = filter.RetryRejectedCall(IntPtr.Zero, 0, SERVERCALL_RETRYLATER);

            Assert.AreEqual(100, first, "First backoff should be the initial BackoffStepMilliseconds.");
            Assert.AreEqual(200, second, "Second backoff should be double the first.");
            Assert.AreEqual(400, third, "Third backoff should be double the second.");

            // Drive the filter until it cancels. After the budget is exhausted it must return -1.
            int sawCancel = 0;
            for (int i = 0; i < 200; i++)
            {
                int wait = filter.RetryRejectedCall(IntPtr.Zero, 0, SERVERCALL_RETRYLATER);
                if (wait == -1)
                {
                    sawCancel = 1;
                    break;
                }
                Assert.IsTrue(wait <= 2000, "Each individual retry must be capped at 2000 ms; got " + wait);
            }
            Assert.AreEqual(1, sawCancel, "Filter must eventually return -1 once MaxRetryMilliseconds is exceeded.");
        }

        [TestMethod]
        public void OneNoteMessageFilter_RetryRejectedCall_CancelOnUnknownReject()
        {
            // dwRejectType values outside the documented SERVERCALL_REJECTED (0) /
            // SERVERCALL_RETRYLATER (2) set must cause an immediate -1 (cancel).
            var filter = new OneNoteMessageFilter();

            Assert.AreEqual(-1, filter.RetryRejectedCall(IntPtr.Zero, 0, 1));
            Assert.AreEqual(-1, filter.RetryRejectedCall(IntPtr.Zero, 0, 7));
            Assert.AreEqual(-1, filter.RetryRejectedCall(IntPtr.Zero, 0, 99));
        }

        [TestMethod]
        public void OneNoteMessageFilter_RetryRejectedCall_ShutdownShortCircuits()
        {
            const int SERVERCALL_RETRYLATER = 2;
            var filter = new OneNoteMessageFilter();

            // Sanity: before shutdown, a retryable reject yields a positive backoff.
            int beforeShutdown = filter.RetryRejectedCall(IntPtr.Zero, 0, SERVERCALL_RETRYLATER);
            Assert.IsTrue(beforeShutdown > 0, "Pre-shutdown should yield a positive backoff, got " + beforeShutdown);

            filter.SignalShutdown();

            // After shutdown the filter must return -1 regardless of dwRejectType.
            Assert.AreEqual(-1, filter.RetryRejectedCall(IntPtr.Zero, 0, SERVERCALL_RETRYLATER));
            Assert.AreEqual(-1, filter.RetryRejectedCall(IntPtr.Zero, 0, 0));
            Assert.AreEqual(-1, filter.RetryRejectedCall(IntPtr.Zero, 0, 42));
        }

        [TestMethod]
        public void StaWorker_PostedActionRunsOnStaThread()
        {
            // Verifies the worker actually pumps queued actions on an STA thread. Completion is
            // signalled via ManualResetEventSlim - never Thread.Sleep - so the test is deterministic.
            // The exception sink is a no-op so a failure inside the action does not pop a MessageBox
            // from the default sink.
            var filter = new OneNoteMessageFilter();
            var worker = new StaWorker("UnitTest-Apartment", filter, _ => { });
            try
            {
                worker.Start();

                ApartmentState observed = ApartmentState.Unknown;
                using (var done = new ManualResetEventSlim(false))
                {
                    worker.Post(() =>
                    {
                        observed = Thread.CurrentThread.GetApartmentState();
                        done.Set();
                    });

                    Assert.IsTrue(done.Wait(2000), "Posted action did not run within 2 s.");
                }

                Assert.AreEqual(ApartmentState.STA, observed, "Worker thread must run posted actions on an STA apartment.");
            }
            finally
            {
                worker.Stop(TimeSpan.FromSeconds(2));
            }
        }

        [TestMethod]
        public void StaWorker_StopJoinsCleanly()
        {
            // Constructing a worker, starting it, and immediately stopping it must drain and join
            // within the timeout (no actions posted, so the consuming loop exits as soon as
            // CompleteAdding flips the queue).
            var filter = new OneNoteMessageFilter();
            var worker = new StaWorker("UnitTest-StopClean", filter, _ => { });
            worker.Start();

            bool joined = worker.Stop(TimeSpan.FromSeconds(2));

            Assert.IsTrue(joined, "Worker must join cleanly within the 2 s timeout when no work is queued.");
        }

        [TestMethod]
        public void StaWorker_ExceptionInActionDoesNotKillThread()
        {
            // Per Step 2 spec: a per-action try/catch keeps the worker alive across non-fatal
            // exceptions. Post a throwing action then a signalling action and assert the second
            // one still runs. The exception sink swallows so the test output stays clean.
            var filter = new OneNoteMessageFilter();
            var worker = new StaWorker("UnitTest-SurviveThrow", filter, _ => { });
            try
            {
                worker.Start();

                using (var done = new ManualResetEventSlim(false))
                {
                    worker.Post(() => { throw new InvalidOperationException("boom"); });
                    worker.Post(() => done.Set());

                    Assert.IsTrue(done.Wait(2000), "Worker must continue processing after a non-fatal exception in a previous action.");
                }
            }
            finally
            {
                worker.Stop(TimeSpan.FromSeconds(2));
            }
        }

        [TestMethod]
        public void EditFormattedCode_OutlineSelected()
        {
            string htmlCode = Resource1.HTMLContent4;
            string[] pos = null;
            HighLightParameter param = new HighLightParameter();
            param.ShowLineNumber = true;
            param.HighlightColor = System.Drawing.Color.FromArgb(240, 240, 240);


            AddIn addIn = new AddIn();
            addIn.ns = @"http://schemas.microsoft.com/office/onenote/2013/onenote";

            var outline = XDocument.Parse(Resource1.Page4).Descendants(addIn.ns + "Outline")
                                   .Where(n => n.Attribute("selected") != null && (n.Attribute("selected").Value == "all" || n.Attribute("selected").Value == "partial"))
                                   .FirstOrDefault();

            //Arrange
            ExeConfigurationFileMap fileMap = new ExeConfigurationFileMap { ExeConfigFilename = "Test.config" };
            Configuration configuration = ConfigurationManager.OpenMappedExeConfiguration(fileMap, ConfigurationUserLevel.None);

            HighLightSection config = configuration.GetSection("HighLightSection") as HighLightSection;

            bool selectedTextFormated = false;
            var page4Root = XDocument.Parse(Resource1.Page4).Root;
            addIn.GetSelectedText(page4Root, out selectedTextFormated);


            XDocument output = addIn.InsertHighLightCode(htmlCode, pos, param, outline, config, selectedTextFormated, addIn.IsSelectedTextInline(page4Root));

            Assert.AreEqual(Resource1.Output4, output.ToString(), false);
        }
    }

    /// <summary>
    /// Tests covering the dynamic-menu payload produced by
    /// <c>AddIn.BuildMoreLanguagesMenuXml</c>. Phase 2 of the language picker plan
    /// (plan section 9.1 "RibbonInvalidationTests"): verify that the returned
    /// fragment is well-formed and declares the customUI namespace on its root
    /// &lt;menu&gt; element, that pinned entries are excluded, and that items are
    /// sorted by display label.
    /// </summary>
    [TestClass]
    public class RibbonInvalidationTests
    {
        private static LanguageDescriptor Desc(string tag, string label, string image = "Other.png", string screentip = null)
        {
            return new LanguageDescriptor("button_" + tag, tag, label, screentip, image, null);
        }

        private static LanguageSettings BuildSettings(IReadOnlyList<string> pinned, IReadOnlyList<string> enabled)
        {
            // Drive the public API to construct an instance with the desired contents.
            // LoadOrSeed has private ctors but it accepts a defaultPinned argument, so we
            // seed via a missing path then mutate to add any extra enabled tags.
            var seed = LanguageSettings.LoadOrSeed(
                path: null,
                knownTags: null,
                defaultPinned: pinned);
            foreach (var tag in enabled)
            {
                seed.Enable(tag);
            }
            return seed;
        }

        [TestMethod]
        public void GetMoreLanguagesMenu_RootDeclaresCustomUiNamespace()
        {
            var registry = new[]
            {
                Desc("cs", "C#"),
                Desc("go", "Go"),
            };
            var settings = BuildSettings(pinned: new[] { "cs" }, enabled: new[] { "cs", "go" });

            string xml = AddIn.BuildMoreLanguagesMenuXml(settings, registry);

            // Parse and assert the root carries the expected xmlns.
            var doc = XDocument.Parse(xml);
            Assert.AreEqual("menu", doc.Root.Name.LocalName);
            Assert.AreEqual(AddIn.CustomUiNamespace, doc.Root.Name.NamespaceName,
                "Dynamic-menu root must declare the customUI xmlns or Office silently drops the menu.");
        }

        [TestMethod]
        public void GetMoreLanguagesMenu_OmitsPinnedAndDisabledLanguages()
        {
            var registry = new[]
            {
                Desc("cs", "C#"),
                Desc("go", "Go"),
                Desc("ts", "TypeScript"),
                Desc("rs", "Rust"),
            };
            // pinned: cs ; enabled (full set incl. pinned): cs, go, ts ; rs is disabled.
            var settings = BuildSettings(pinned: new[] { "cs" }, enabled: new[] { "cs", "go", "ts" });

            string xml = AddIn.BuildMoreLanguagesMenuXml(settings, registry);
            var doc = XDocument.Parse(xml);
            XNamespace ns = AddIn.CustomUiNamespace;

            var tags = doc.Root.Elements(ns + "button")
                                .Select(b => (string)b.Attribute("tag"))
                                .ToList();

            CollectionAssert.AreEquivalent(new[] { "go", "ts" }, tags,
                "Dynamic menu must list enabled-but-not-pinned tags only.");
            Assert.IsFalse(tags.Contains("cs"), "Pinned tags must not appear in the dynamic menu.");
            Assert.IsFalse(tags.Contains("rs"), "Disabled tags must not appear in the dynamic menu.");
        }

        [TestMethod]
        public void GetMoreLanguagesMenu_ItemsAreSortedByLabel()
        {
            var registry = new[]
            {
                Desc("ts", "TypeScript"),
                Desc("go", "Go"),
                Desc("rs", "Rust"),
                Desc("py", "Python"),
            };
            var settings = BuildSettings(pinned: new string[0], enabled: new[] { "ts", "go", "rs", "py" });

            string xml = AddIn.BuildMoreLanguagesMenuXml(settings, registry);
            var doc = XDocument.Parse(xml);
            XNamespace ns = AddIn.CustomUiNamespace;

            var labels = doc.Root.Elements(ns + "button")
                                 .Select(b => (string)b.Attribute("label"))
                                 .ToList();

            CollectionAssert.AreEqual(new[] { "Go", "Python", "Rust", "TypeScript" }, labels,
                "Dynamic-menu items must be sorted by display label.");
        }

        [TestMethod]
        public void GetMoreLanguagesMenu_EachItemReusesAddInButtonClicked()
        {
            var registry = new[] { Desc("go", "Go"), Desc("ts", "TypeScript") };
            var settings = BuildSettings(pinned: new string[0], enabled: new[] { "go", "ts" });

            string xml = AddIn.BuildMoreLanguagesMenuXml(settings, registry);
            var doc = XDocument.Parse(xml);
            XNamespace ns = AddIn.CustomUiNamespace;

            foreach (var btn in doc.Root.Elements(ns + "button"))
            {
                Assert.AreEqual("AddInButtonClicked", (string)btn.Attribute("onAction"),
                    "Dynamic-menu buttons must reuse the existing onAction so the highlight path is shared.");
                Assert.IsTrue(((string)btn.Attribute("id")).StartsWith("dyn_", StringComparison.Ordinal),
                    "Dynamic-menu button ids must be namespaced ('dyn_<tag>') so they cannot collide with pinned button ids.");
                Assert.IsFalse(string.IsNullOrEmpty((string)btn.Attribute("image")),
                    "Every dynamic-menu button must carry an image attribute.");
            }
        }

        [TestMethod]
        public void GetMoreLanguagesMenu_EmptyWhenNothingEnabledButPinned()
        {
            var registry = new[] { Desc("cs", "C#") };
            var settings = BuildSettings(pinned: new[] { "cs" }, enabled: new[] { "cs" });

            string xml = AddIn.BuildMoreLanguagesMenuXml(settings, registry);
            var doc = XDocument.Parse(xml);
            XNamespace ns = AddIn.CustomUiNamespace;

            Assert.AreEqual("menu", doc.Root.Name.LocalName);
            Assert.AreEqual(AddIn.CustomUiNamespace, doc.Root.Name.NamespaceName);
            Assert.AreEqual(0, doc.Root.Elements(ns + "button").Count(),
                "Menu must be empty when every enabled tag is also pinned.");
        }

        [TestMethod]
        public void GetMoreLanguagesMenu_NullSettingsYieldsWellFormedEmptyMenu()
        {
            var registry = new[] { Desc("cs", "C#") };

            string xml = AddIn.BuildMoreLanguagesMenuXml(settings: null, registry: registry);
            var doc = XDocument.Parse(xml);

            Assert.AreEqual("menu", doc.Root.Name.LocalName);
            Assert.AreEqual(AddIn.CustomUiNamespace, doc.Root.Name.NamespaceName);
            Assert.IsFalse(doc.Root.HasElements);
        }

        [TestMethod]
        public void GetMoreLanguagesMenu_TagOutsideRegistryIsDropped()
        {
            // A tag persisted in languages.json that no longer appears in ribbon.xml
            // must not be rendered (we have no descriptor to draw the button from).
            var registry = new[] { Desc("cs", "C#") };
            var settings = BuildSettings(pinned: new string[0], enabled: new[] { "ghost-tag" });

            string xml = AddIn.BuildMoreLanguagesMenuXml(settings, registry);
            var doc = XDocument.Parse(xml);
            XNamespace ns = AddIn.CustomUiNamespace;

            Assert.AreEqual(0, doc.Root.Elements(ns + "button").Count(),
                "Tags without a matching descriptor must be silently dropped from the dynamic menu.");
        }

        [TestMethod]
        public void GetMoreLanguagesMenu_IncludesEnabledNotPinnedNewLanguage()
        {
            // Plan section 5.4: round-trip one of the six newly-added languages
            // through the registry + dynamic-menu builder. The user enables YAML
            // via the Languages... dialog but does not pin it; the dynamic menu
            // must therefore include a dyn_yaml button carrying tag="yaml".
            var registry = new[]
            {
                Desc("cs", "C#"),
                Desc("yaml", "YAML"),
            };
            var settings = BuildSettings(pinned: new string[0], enabled: new[] { "yaml" });

            Assert.IsTrue(settings.IsEnabled("yaml"));
            Assert.IsFalse(settings.IsPinned("yaml"));
            Assert.IsTrue(settings.IsVisibleInMoreMenu("yaml"));

            string xml = AddIn.BuildMoreLanguagesMenuXml(settings, registry);
            var doc = XDocument.Parse(xml);
            XNamespace ns = AddIn.CustomUiNamespace;

            Assert.AreEqual("menu", doc.Root.Name.LocalName);
            Assert.AreEqual(AddIn.CustomUiNamespace, doc.Root.Name.NamespaceName);

            var yamlBtn = doc.Root.Elements(ns + "button")
                                  .FirstOrDefault(e => (string)e.Attribute("tag") == "yaml");
            Assert.IsNotNull(yamlBtn, "Dynamic menu must include the enabled-not-pinned 'yaml' tag.");
            Assert.AreEqual("dyn_yaml", (string)yamlBtn.Attribute("id"));
            Assert.AreEqual("YAML", (string)yamlBtn.Attribute("label"));
            Assert.AreEqual("AddInButtonClicked", (string)yamlBtn.Attribute("onAction"));
        }

        [TestMethod]
        public void GetMoreLanguagesMenu_LabelWithSpecialCharsIsEscaped()
        {
            // XmlWriter must escape & < > " in attribute values - this is the whole reason
            // the dynamic-menu fragment is built via XmlWriter rather than string concatenation.
            var registry = new[] { Desc("ab", "A & B <c>") };
            var settings = BuildSettings(pinned: new string[0], enabled: new[] { "ab" });

            string xml = AddIn.BuildMoreLanguagesMenuXml(settings, registry);
            var doc = XDocument.Parse(xml);  // would throw if escaping were wrong
            XNamespace ns = AddIn.CustomUiNamespace;

            var btn = doc.Root.Element(ns + "button");
            Assert.IsNotNull(btn);
            Assert.AreEqual("A & B <c>", (string)btn.Attribute("label"));
        }
    }
}
