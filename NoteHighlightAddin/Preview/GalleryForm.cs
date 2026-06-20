using GenerateHighlightContent;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NoteHighlightAddin.Preview
{
    /// <summary>
    /// Modal dialog that renders the current editor code across every installed
    /// theme as clickable tiles in one <see cref="WebBrowser"/>, reusing the live
    /// pane's render path so each tile matches its single-pane preview. The pass runs
    /// background-threaded with bounded concurrency and is cancelled before teardown
    /// so closing mid-render orphans no highlight.exe child. A tile click round-trips
    /// the selection through the host dropdown; the gallery never persists the setting.
    /// </summary>
    public partial class GalleryForm : Form
    {
        private const int MaxConcurrentRenders = 8;
        private const string SelectScheme = "nhtheme";

        private readonly SynchronizationContext _uiContext;
        private readonly MainForm _owner;
        private readonly bool _darkMode;
        private readonly System.Drawing.Color _boxColor;
        private readonly bool _showTableBorder;
        private readonly ComboBox _styleDropdown;

        private readonly CancellationTokenSource _renderCts = new CancellationTokenSource();
        private int _completedCount;
        private bool _torndown;

        /// <param name="owner">Host form; supplies <c>BuildPreviewParameter</c> so each tile is built from identical state.</param>
        /// <param name="styleDropdown">Host style combo; the tile-click round-trip sets its SelectedIndex so existing wiring persists the choice.</param>
        public GalleryForm(MainForm owner, ComboBox styleDropdown, bool darkMode, System.Drawing.Color boxColor, bool showTableBorder)
        {
            _owner = owner ?? throw new ArgumentNullException(nameof(owner));
            _styleDropdown = styleDropdown ?? throw new ArgumentNullException(nameof(styleDropdown));
            _darkMode = darkMode;
            _boxColor = boxColor;
            _showTableBorder = showTableBorder;

            InitializeComponent();

            _uiContext = SynchronizationContext.Current ?? new SynchronizationContext();

            this.browser.Navigating += Browser_Navigating;
            this.Shown += GalleryForm_Shown;
            this.FormClosing += GalleryForm_FormClosing;
        }

        private void GalleryForm_Shown(object sender, EventArgs e)
        {
            StartRenderPass();
        }

        private void StartRenderPass()
        {
            string addinDir = AddIn.GetAddinDirectory();
            if (string.IsNullOrEmpty(addinDir))
            {
                ShowFatal("Could not resolve the add-in install directory.");
                return;
            }

            List<string> themes = EnumerateThemes(addinDir);
            if (themes.Count == 0)
            {
                ShowFatal("No highlight themes were found next to the add-in.");
                return;
            }

            // Snapshot parameters on the UI thread; BuildPreviewParameter reads form controls the workers must not touch.
            var parameters = new HighLightParameter[themes.Count];
            for (int i = 0; i < themes.Count; i++)
            {
                parameters[i] = _owner.BuildPreviewParameter(themes[i]);
            }

            this.progressBar.Minimum = 0;
            this.progressBar.Maximum = themes.Count;
            this.progressBar.Value = 0;
            UpdateProgressLabel(0, themes.Count);

            CancellationToken token = _renderCts.Token;
            Task.Run(() => RenderAllAsync(addinDir, themes, parameters, token));
        }

        // Same basenames as MainForm.LoadThemes, or the name-to-index lookup misses.
        private static List<string> EnumerateThemes(string addinDir)
        {
            try
            {
                var generate = new GenerateHighLight(addinDir);
                HighLightSection section = generate.Config;
                string themesDir = Path.Combine(Path.Combine(addinDir, section.FolderName), section.ThemeFolder);
                if (!Directory.Exists(themesDir)) return new List<string>();

                return Directory.GetFiles(themesDir, "*.theme")
                    .Select(Path.GetFileNameWithoutExtension)
                    .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            catch
            {
                return new List<string>();
            }
        }

        private async Task RenderAllAsync(string addinDir, List<string> themes, HighLightParameter[] parameters, CancellationToken token)
        {
            var tiles = new string[themes.Count];
            using (var gate = new SemaphoreSlim(MaxConcurrentRenders))
            {
                var tasks = new List<Task>(themes.Count);
                for (int i = 0; i < themes.Count; i++)
                {
                    int index = i;
                    string theme = themes[index];
                    HighLightParameter parameter = parameters[index];
                    tasks.Add(Task.Run(async () =>
                    {
                        await gate.WaitAsync(token).ConfigureAwait(false);
                        try
                        {
                            token.ThrowIfCancellationRequested();
                            tiles[index] = RenderTile(addinDir, theme, parameter, token);
                        }
                        finally
                        {
                            gate.Release();
                            ReportProgress(themes.Count);
                        }
                    }, token));
                }

                try
                {
                    await Task.WhenAll(tasks).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    return;
                }
                catch
                {
                    // RenderTile captures per-tile failures; the partial document still renders.
                }
            }

            if (token.IsCancellationRequested) return;

            string document = BuildGalleryDocument(themes, tiles);
            _uiContext.Post(_ =>
            {
                if (token.IsCancellationRequested) return;
                if (IsDisposed || _torndown) return;
                try { browser.DocumentText = document; } catch { }
                this.progressPanel.Visible = false;
            }, null);
        }

        private string RenderTile(string addinDir, string theme, HighLightParameter parameter, CancellationToken token)
        {
            try
            {
                // Each tile owns its parameter instance, so this unique FileName cannot race siblings.
                parameter.FileName = "gallery-" + Guid.NewGuid().ToString("N");

                var generate = new GenerateHighLight(addinDir);
                string outputPath = generate.GenerateHighLightCode(parameter, token);

                string rawHtml = File.ReadAllText(outputPath, new UTF8Encoding(false));
                try { File.Delete(outputPath); } catch { }

                string wrapped = PreviewHtmlWrapper.Wrap(rawHtml, _darkMode, _boxColor, _showTableBorder);
                return WrapTile(theme, wrapped, false);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch
            {
                return WrapTile(theme, "<div class=\"nh-tile-error\">render failed</div>", true);
            }
        }

        private void ReportProgress(int total)
        {
            int done = Interlocked.Increment(ref _completedCount);
            _uiContext.Post(_ =>
            {
                if (IsDisposed || _torndown) return;
                if (progressBar.Value < done && done <= progressBar.Maximum) progressBar.Value = done;
                UpdateProgressLabel(done, total);
            }, null);
        }

        private void UpdateProgressLabel(int done, int total)
        {
            lblProgress.Text = string.Format(CultureInfo.CurrentCulture, "Rendering themes... {0} / {1}", done, total);
        }

        // Anchor carries the URL-encoded theme on a custom scheme for Browser_Navigating to round-trip.
        private static string WrapTile(string theme, string innerHtml, bool failed)
        {
            string encoded = Uri.EscapeDataString(theme);
            string label = HtmlEncode(theme);
            var sb = new StringBuilder();
            sb.Append("<div class=\"nh-tile").Append(failed ? " nh-tile-failed" : string.Empty).Append("\">");
            sb.Append("<a class=\"nh-tile-link\" href=\"").Append(SelectScheme).Append(':').Append(encoded).Append("\">");
            sb.Append("<div class=\"nh-tile-caption\">").Append(label).Append("</div>");
            sb.Append("<div class=\"nh-tile-body\">").Append(innerHtml).Append("</div>");
            sb.Append("</a></div>");
            return sb.ToString();
        }

        private string BuildGalleryDocument(List<string> themes, string[] tiles)
        {
            var sb = new StringBuilder();
            sb.Append("<!DOCTYPE html><html><head>");
            sb.Append("<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\">");
            sb.Append("<style>");
            sb.Append("html,body{margin:0;padding:12px;font-family:Segoe UI,Tahoma,sans-serif;");
            sb.Append(_darkMode ? "background:#1e1e1e;color:#ddd;" : "background:#f5f5f5;color:#222;");
            sb.Append("}");
            sb.Append(".nh-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));grid-gap:12px;}");
            sb.Append(".nh-tile{border:1px solid #999;border-radius:4px;overflow:hidden;background:#fff;}");
            sb.Append(".nh-tile-failed{border-color:#c0392b;}");
            sb.Append(".nh-tile-link{text-decoration:none;color:inherit;display:block;cursor:pointer;}");
            sb.Append(".nh-tile-caption{padding:4px 8px;font-size:12px;font-weight:bold;background:#e8e8e8;color:#222;border-bottom:1px solid #ccc;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}");
            sb.Append(".nh-tile-body{max-height:200px;overflow:hidden;}");
            sb.Append(".nh-tile-body pre{white-space:pre !important;word-wrap:normal !important;overflow-wrap:normal !important;}");
            sb.Append(".nh-tile-error{padding:16px;color:#c0392b;font-size:12px;}");
            sb.Append("</style></head><body>");
            sb.Append("<div class=\"nh-grid\">");
            for (int i = 0; i < tiles.Length; i++)
            {
                sb.Append(tiles[i] ?? WrapTile(themes[i], "<div class=\"nh-tile-error\">not rendered</div>", true));
            }
            sb.Append("</div></body></html>");
            return sb.ToString();
        }

        private void Browser_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            string url = e.Url != null ? e.Url.OriginalString : null;
            string prefix = SelectScheme + ":";
            if (string.IsNullOrEmpty(url) || !url.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            // Never let MSHTML actually navigate away from the gallery document.
            e.Cancel = true;

            string theme;
            try { theme = Uri.UnescapeDataString(url.Substring(prefix.Length)); }
            catch { return; }

            int index = FindThemeIndex(theme);
            if (index < 0) return;

            _renderCts.Cancel();
            _styleDropdown.SelectedIndex = index;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private int FindThemeIndex(string theme)
        {
            for (int i = 0; i < _styleDropdown.Items.Count; i++)
            {
                if (string.Equals(Convert.ToString(_styleDropdown.Items[i]), theme, StringComparison.Ordinal))
                {
                    return i;
                }
            }
            return -1;
        }

        private void ShowFatal(string message)
        {
            lblProgress.Text = message;
            progressBar.Visible = false;
        }

        private static string HtmlEncode(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            var sb = new StringBuilder(s.Length);
            foreach (char ch in s)
            {
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

        private void GalleryForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            TeardownBrowser();
        }

        private void TeardownBrowser()
        {
            if (_torndown) return;
            _torndown = true;

            // Cancel before teardown so no further highlight.exe children launch and in-flight ones die.
            try { _renderCts.Cancel(); } catch (ObjectDisposedException) { }

            if (browser != null)
            {
                // Canonical MSHTML teardown (mirrors PreviewPane.Dispose); never Marshal.ReleaseComObject the browser - AxHost owns it, double-release AVs.
                try { browser.Stop(); } catch { }
                try { browser.Navigate("about:blank"); } catch { }
                try { Application.DoEvents(); } catch { }
                try { browser.Dispose(); } catch { }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                TeardownBrowser();

                try { _renderCts.Dispose(); } catch { }

                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }
    }
}
