using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NoteHighlightAddin.Preview;

namespace UnitTesting
{
    [TestClass]
    public class PreviewHtmlWrapperTests
    {
        [TestMethod]
        public void Wrap_OutputContainsIEEdgeMetaTag()
        {
            string output = PreviewHtmlWrapper.Wrap("<html><body><pre>hi</pre></body></html>", false, null);

            Assert.IsTrue(
                output.IndexOf("<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\">", StringComparison.Ordinal) >= 0,
                "Wrap output must contain the IE=edge X-UA-Compatible meta tag.");
        }

        [TestMethod]
        public void Wrap_DarkMode_StripsPreBackgroundColor()
        {
            string input = "<pre style=\"background-color:#fff; color:#000;\">code</pre>";

            string output = PreviewHtmlWrapper.Wrap(input, true, null);

            Assert.IsFalse(
                output.IndexOf("background-color", StringComparison.Ordinal) >= 0,
                "Dark mode must strip background-color from the <pre> tag.");
        }

        [TestMethod]
        public void Wrap_LightMode_PreservesPreBackgroundColor()
        {
            string input = "<pre style=\"background-color:#fff; color:#000;\">code</pre>";

            string output = PreviewHtmlWrapper.Wrap(input, false, null);

            Assert.IsTrue(
                output.IndexOf("background-color", StringComparison.Ordinal) >= 0,
                "Light mode must preserve background-color on the <pre> tag.");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void StripPreBackgroundColor_ThrowsOnMissingSemicolon()
        {
            PreviewHtmlWrapper.StripPreBackgroundColor("<pre style=\"background-color:#fff\">");
        }

        [TestMethod]
        public void Wrap_StripsXmlProlog()
        {
            string input = "<?xml version=\"1.0\" encoding=\"utf-8\"?><html><body><pre>hi</pre></body></html>";

            string output = PreviewHtmlWrapper.Wrap(input, false, null);

            Assert.IsFalse(
                output.IndexOf("<?xml", StringComparison.Ordinal) >= 0,
                "Wrap must strip a leading <?xml ?> prolog so MSHTML stays in HTML mode.");
        }
    }
}
