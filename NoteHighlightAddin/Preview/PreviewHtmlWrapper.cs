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
        /// </summary>
        public static string Wrap(string highlightOutputHtml, bool darkMode, Color? formBackground)
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
            if (formBackground.HasValue)
            {
                Color c = formBackground.Value;
                string hex = string.Format("#{0:X2}{1:X2}{2:X2}", c.R, c.G, c.B);
                sb.Append("body { background:").Append(hex).Append("; }");
            }
            sb.Append("</style></head><body>");
            sb.Append(html);
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
    }
}
