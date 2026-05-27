using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;

namespace NoteHighlightAddin.Preview
{
    /// <summary>
    /// Helpers for adapting <c>highlight.exe</c> HTML output for in-process
    /// preview hosting in an MSHTML <see cref="System.Windows.Forms.WebBrowser"/>.
    /// </summary>
    public static class PreviewHtmlWrapper
    {
        private static readonly Regex XmlPrologRegex = new Regex(
            @"^\s*<\?xml[^>]*\?>",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Wraps raw <c>highlight.exe</c> HTML into a self-contained document
        /// suitable for hosting in an MSHTML <see cref="System.Windows.Forms.WebBrowser"/>.
        /// Strips any leading <c>&lt;?xml ?&gt;</c> prolog so MSHTML stays in HTML
        /// (not XML) mode under <c>IE=edge</c>, and optionally strips the
        /// <c>background-color</c> declaration from the first <c>&lt;pre&gt;</c>
        /// tag when <paramref name="darkMode"/> is true so dark themes render
        /// with correct contrast.
        /// <para>
        /// To mirror the real OneNote insertion, <paramref name="boxColor"/> is
        /// applied as the background of the container that holds the code block
        /// (the equivalent of the OneNote cell <c>shadingColor</c>), so it shows
        /// as the surround around the <c>&lt;pre&gt;</c> rather than over the code
        /// text - the <c>&lt;pre&gt;</c> keeps its own theme background on top. A
        /// null <paramref name="boxColor"/>, or one whose alpha byte is 0
        /// (transparent, the OneNote <c>"none"</c> case), leaves the default white
        /// surround. When <paramref name="showTableBorder"/> is true a 1px solid
        /// neutral grey border is drawn around that container, matching OneNote's
        /// default cell border.
        /// </para>
        /// </summary>
        public static string Wrap(string highlightOutputHtml, bool darkMode, Color? boxColor, bool showTableBorder)
        {
            string html = XmlPrologRegex.Replace(highlightOutputHtml ?? string.Empty, string.Empty);

            if (darkMode)
            {
                int preStart = html.IndexOf("<pre");
                if (preStart >= 0)
                {
                    int preEnd = html.IndexOf('>', preStart);
                    if (preEnd >= 0)
                    {
                        int length = preEnd - preStart + 1;
                        string preTag = html.Substring(preStart, length);
                        string stripped = StripPreBackgroundColor(preTag);
                        html = html.Remove(preStart, length).Insert(preStart, stripped);
                    }
                }
            }

            StringBuilder sb = new StringBuilder();
            sb.Append("<!DOCTYPE html><html><head>");
            sb.Append("<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\">");
            sb.Append("<style>html,body { margin:0; padding:8px; overflow:auto; }");
            // Force the code block to never reflow: long lines scroll horizontally
            // (html,body overflow:auto supplies the scrollbar) so the preview mirrors
            // how the code is inserted rather than word-wrapping. highlight.exe with
            // --inline-css emits the <pre> with an inline "white-space: pre-wrap;", and
            // an inline declaration outranks a plain type-selector rule - so !important
            // is required here to override it (an important author rule beats a normal
            // inline declaration). It does not depend on highlight.exe's exact spacing,
            // unlike rewriting the inline string would.
            sb.Append("pre { white-space:pre !important; word-wrap:normal !important; overflow-wrap:normal !important; }");

            sb.Append(".nh-preview-box { padding:4px;");
            if (boxColor.HasValue && boxColor.Value.A != 0)
            {
                Color c = boxColor.Value;
                string hex = string.Format("#{0:X2}{1:X2}{2:X2}", c.R, c.G, c.B);
                sb.Append(" background:").Append(hex).Append(";");
            }
            if (showTableBorder)
            {
                sb.Append(" border:1px solid #A6A6A6;");
            }
            sb.Append(" }");

            sb.Append("</style></head><body>");
            sb.Append("<div class=\"nh-preview-box\">");
            sb.Append(html);
            sb.Append("</div>");
            sb.Append("</body></html>");
            return sb.ToString();
        }

        /// <summary>
        /// Removes the <c>background-color: ...;</c> declaration from a single
        /// <c>&lt;pre&gt;</c> opening tag, mirroring the existing in-line
        /// implementations at <c>AddIn.cs:843</c> and <c>MainForm.cs:195</c>.
        /// Returns <paramref name="preTag"/> unchanged when no
        /// <c>background-color</c> substring is present.
        /// </summary>
        /// <remarks>
        /// Throws <see cref="System.ArgumentOutOfRangeException"/> when
        /// <c>background-color</c> is present but has no terminating <c>;</c>
        /// (because <c>IndexOf(';', bcIndex)</c> returns <c>-1</c> and the
        /// computed length passed to <see cref="string.Remove(int, int)"/> goes
        /// negative). This behaviour is intentionally preserved from the two
        /// original call sites - <c>highlight.exe</c> reliably emits the
        /// terminating semicolon, and silently changing the semantics during a
        /// refactor was judged riskier than keeping them.
        /// </remarks>
        public static string StripPreBackgroundColor(string preTag)
        {
            int bcIndex = preTag.IndexOf("background-color");
            if (bcIndex < 0)
            {
                return preTag;
            }
            int semicolonIndex = preTag.IndexOf(';', bcIndex);
            return preTag.Remove(bcIndex, semicolonIndex - bcIndex + 1);
        }

        /// <summary>
        /// Builds a self-contained error document for display in the preview
        /// <see cref="System.Windows.Forms.WebBrowser"/> when <c>highlight.exe</c>
        /// fails. The exception message (which embeds the captured stderr
        /// snippet for <c>InvalidOperationException</c> thrown by
        /// <c>ProcessHelper</c>) and, if present, the inner exception's
        /// message are rendered inside a plain <c>&lt;pre&gt;</c> block.
        /// Both messages are HTML-escaped so that stderr containing
        /// <c>&lt;</c>, <c>&gt;</c>, or <c>&amp;</c> cannot inject markup.
        /// Mirrors the meta/header structure of <see cref="Wrap"/> so the
        /// document still pins MSHTML to <c>IE=edge</c>.
        /// </summary>
        public static string BuildErrorDocument(System.Exception exception, Color? formBackground)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<!DOCTYPE html><html><head>");
            sb.Append("<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\">");
            sb.Append("<style>html,body { margin:0; padding:8px; overflow:auto; font-family:Segoe UI,Tahoma,sans-serif; }");
            if (formBackground.HasValue)
            {
                Color c = formBackground.Value;
                string hex = string.Format("#{0:X2}{1:X2}{2:X2}", c.R, c.G, c.B);
                sb.Append("body { background:").Append(hex).Append("; }");
            }
            sb.Append("h2 { color:#b00020; font-size:13px; margin:0 0 6px 0; }");
            sb.Append("pre { white-space:pre-wrap; word-break:break-word; background:#fff3f3; border:1px solid #f0c0c0; padding:8px; color:#3a0000; font-family:Consolas,monospace; font-size:12px; }");
            sb.Append("</style></head><body>");
            sb.Append("<h2>Preview render failed</h2>");
            sb.Append("<pre>");

            if (exception == null)
            {
                sb.Append("(unknown error)");
            }
            else
            {
                sb.Append(HtmlEscape(exception.Message ?? string.Empty));
                if (exception.InnerException != null)
                {
                    sb.Append("\n\nInner: ");
                    sb.Append(HtmlEscape(exception.InnerException.Message ?? string.Empty));
                }
            }

            sb.Append("</pre>");
            sb.Append("</body></html>");
            return sb.ToString();
        }

        private static string HtmlEscape(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            StringBuilder sb = new StringBuilder(s.Length);
            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];
                switch (ch)
                {
                    case '&': sb.Append("&amp;"); break;
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    case '"': sb.Append("&quot;"); break;
                    case '\'': sb.Append("&#39;"); break;
                    default: sb.Append(ch); break;
                }
            }
            return sb.ToString();
        }
    }
}
