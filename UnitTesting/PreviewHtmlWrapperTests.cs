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

        [TestMethod]
        public void BuildErrorDocument_EscapesStderrMarkup()
        {
            // ProcessHelper.InvalidOperationException embeds the stderr snippet
            // in the message verbatim. Stderr can contain raw <, >, &; the
            // error document must not let that inject markup into MSHTML.
            var ex = new InvalidOperationException(
                "Process 'highlight.exe' exited with code 1. Stderr: <script>alert(1)</script> & some & more");

            string output = PreviewHtmlWrapper.BuildErrorDocument(ex, null);

            Assert.IsTrue(output.IndexOf("<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\">", StringComparison.Ordinal) >= 0,
                "Error document must pin MSHTML to IE=edge like Wrap does.");
            Assert.IsFalse(output.IndexOf("<script>alert(1)</script>", StringComparison.Ordinal) >= 0,
                "Raw <script> from stderr must be HTML-escaped, not passed through.");
            Assert.IsTrue(output.IndexOf("&lt;script&gt;", StringComparison.Ordinal) >= 0,
                "Less-than/greater-than from the exception message must be escaped.");
            Assert.IsTrue(output.IndexOf("&amp;", StringComparison.Ordinal) >= 0,
                "Ampersand from the exception message must be escaped.");
        }

        [TestMethod]
        public void BuildErrorDocument_IncludesInnerExceptionMessage()
        {
            var inner = new InvalidOperationException("inner cause text");
            var outer = new InvalidOperationException("outer wrapper text", inner);

            string output = PreviewHtmlWrapper.BuildErrorDocument(outer, null);

            Assert.IsTrue(output.IndexOf("outer wrapper text", StringComparison.Ordinal) >= 0,
                "Outer exception message must appear in the error document.");
            Assert.IsTrue(output.IndexOf("inner cause text", StringComparison.Ordinal) >= 0,
                "Inner exception message must appear in the error document.");
        }

        [TestMethod]
        public void BuildErrorDocument_NullException_DoesNotThrow()
        {
            // Defensive: the catch site should be able to call BuildErrorDocument
            // even on unusual inputs without raising a secondary exception.
            string output = PreviewHtmlWrapper.BuildErrorDocument(null, null);

            Assert.IsTrue(output.IndexOf("<html>", StringComparison.Ordinal) >= 0,
                "Null exception must still produce a wellformed HTML document.");
        }
    }
}
