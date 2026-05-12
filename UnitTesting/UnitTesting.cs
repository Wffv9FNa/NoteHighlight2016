using System;
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
}
