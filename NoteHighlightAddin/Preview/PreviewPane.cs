using GenerateHighlightContent;
using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;

namespace NoteHighlightAddin.Preview
{
    /// <summary>
    /// WinForms control that hosts a <see cref="WebBrowser"/> and renders the
    /// output of <c>highlight.exe</c> as a live preview. Renders are debounced
    /// and gated by a monotonically-increasing sequence number so that a
    /// stale background render cannot overwrite a fresher one. The hosting
    /// form is responsible for calling <see cref="Dispose()"/> so the MSHTML
    /// COM teardown happens on the STA worker that created the control.
    /// </summary>
    public partial class PreviewPane : UserControl
    {
        private const int DebounceMs = 300;

        private readonly SynchronizationContext _uiContext;
        private readonly Timer _debounceTimer;

        private string _sessionTempDir;
        private int _renderSeq;

        private HighLightParameter _pendingParameters;
        private bool _pendingDarkMode;

        public PreviewPane()
        {
            InitializeComponent();

            _uiContext = SynchronizationContext.Current ?? new SynchronizationContext();

            _debounceTimer = new Timer { Interval = DebounceMs };
            _debounceTimer.Tick += DebounceTimer_Tick;
        }

        /// <summary>
        /// When true, <see cref="Render"/> is a no-op. Set by the host
        /// (e.g. <c>MainForm</c> when <c>_quickStyle</c> is active) so we do
        /// not waste cycles rendering a preview the user will never see.
        /// </summary>
        public bool QuickStyleSuppress { get; set; }

        public void Render(HighLightParameter parameters, bool darkMode)
        {
            if (QuickStyleSuppress) return;
            if (parameters == null) return;

            _renderSeq++;
            _pendingParameters = parameters;
            _pendingDarkMode = darkMode;

            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        public void Clear()
        {
            browser.DocumentText = "<html><body></body></html>";
        }

        private string GetOrCreateSessionTempDir()
        {
            if (_sessionTempDir == null)
            {
                _sessionTempDir = Path.Combine(
                    Path.GetTempPath(),
                    "NoteHighlight2016",
                    "preview-" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(_sessionTempDir);
            }
            return _sessionTempDir;
        }

        private void DebounceTimer_Tick(object sender, EventArgs e)
        {
            _debounceTimer.Stop();

            HighLightParameter snapshot = _pendingParameters;
            bool darkMode = _pendingDarkMode;
            int mySeq = _renderSeq;
            if (snapshot == null) return;

            string addinDir = AddIn.GetAddinDirectory();
            if (string.IsNullOrEmpty(addinDir)) return;

            Task.Run(() =>
            {
                string wrapped;
                try
                {
                    string sessionDir = GetOrCreateSessionTempDir();
                    string scratchName = Path.GetRandomFileName();
                    string scratchPath = Path.Combine(sessionDir, scratchName);

                    File.WriteAllText(scratchPath, snapshot.Content ?? string.Empty, Encoding.UTF8);

                    // GenerateHighLight builds its own input/output paths from
                    // Path.GetTempPath() + parameters.FileName, so feed it a
                    // unique FileName for this render. The pane's own scratch
                    // file is kept inside _sessionTempDir purely for the
                    // disposal contract.
                    var paneParameters = new HighLightParameter
                    {
                        Content = snapshot.Content,
                        CodeType = snapshot.CodeType,
                        HighLightStyle = snapshot.HighLightStyle,
                        ShowLineNumber = snapshot.ShowLineNumber,
                        FileName = "preview-" + Guid.NewGuid().ToString("N"),
                        Font = snapshot.Font,
                        FontSize = snapshot.FontSize,
                        HighlightColor = snapshot.HighlightColor
                    };

                    var generate = new GenerateHighLight(addinDir);
                    string outputPath = generate.GenerateHighLightCode(paneParameters);

                    string rawHtml = File.ReadAllText(outputPath, new UTF8Encoding(false));
                    try { File.Delete(outputPath); } catch { }
                    try { File.Delete(scratchPath); } catch { }

                    wrapped = PreviewHtmlWrapper.Wrap(rawHtml, darkMode, BackColor);
                }
                catch
                {
                    // Phase 2.6 will surface errors into the pane. For Phase 1
                    // we silently leave the previous render on screen.
                    return;
                }

                _uiContext.Post(_ =>
                {
                    if (mySeq != _renderSeq) return;
                    if (IsDisposed) return;
                    browser.DocumentText = wrapped;
                }, null);
            });
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_debounceTimer != null)
                {
                    try { _debounceTimer.Stop(); } catch { }
                    try { _debounceTimer.Dispose(); } catch { }
                }

                if (browser != null)
                {
                    // MSHTML teardown order: stop loads, navigate away so the
                    // document is released, pump the message queue once, then
                    // dispose. Do NOT Marshal.ReleaseComObject the underlying
                    // ActiveX site - AxHost owns it and a double-release AVs.
                    try { browser.Stop(); } catch { }
                    try { browser.Navigate("about:blank"); } catch { }
                    try { Application.DoEvents(); } catch { }
                    try { browser.Dispose(); } catch { }
                }

                if (_sessionTempDir != null)
                {
                    try
                    {
                        if (Directory.Exists(_sessionTempDir))
                            Directory.Delete(_sessionTempDir, true);
                    }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }

                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }
    }
}
