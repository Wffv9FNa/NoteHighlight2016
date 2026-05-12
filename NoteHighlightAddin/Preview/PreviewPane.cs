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
        private const int IndicatorDelayMs = 500;

        private readonly SynchronizationContext _uiContext;
        private readonly Timer _debounceTimer;
        private readonly Timer _indicatorTimer;

        private string _sessionTempDir;
        private int _renderSeq;

        private HighLightParameter _pendingParameters;
        private bool _pendingDarkMode;

        /// <summary>
        /// Token source for the currently-scheduled render. Each debounce tick
        /// cancels the previous CTS (killing any highlight.exe child in flight)
        /// and creates a new one. Defence in depth alongside <c>_renderSeq</c>:
        /// the seq-gate discards a stale apply step even when cancellation has
        /// not yet observed the token.
        /// </summary>
        private CancellationTokenSource _renderCts;

        public PreviewPane()
        {
            InitializeComponent();

            _uiContext = SynchronizationContext.Current ?? new SynchronizationContext();

            _debounceTimer = new Timer { Interval = DebounceMs };
            _debounceTimer.Tick += DebounceTimer_Tick;

            // Indicator timer: started when a render is dispatched, fires once
            // after IndicatorDelayMs to reveal the "rendering..." label. The
            // tick handler stops the timer so it does not fire repeatedly; the
            // label is hidden again on apply, cancel, or error.
            _indicatorTimer = new Timer { Interval = IndicatorDelayMs };
            _indicatorTimer.Tick += IndicatorTimer_Tick;
        }

        private void IndicatorTimer_Tick(object sender, EventArgs e)
        {
            _indicatorTimer.Stop();
            if (IsDisposed) return;
            if (lblRendering != null) lblRendering.Visible = true;
        }

        private void StartIndicatorCountdown()
        {
            // Restart the 500 ms countdown. If a render completes faster than
            // this it never appears; if it overruns, the tick handler reveals
            // the label.
            _indicatorTimer.Stop();
            _indicatorTimer.Start();
        }

        private void HideIndicator()
        {
            _indicatorTimer.Stop();
            if (lblRendering != null && lblRendering.Visible) lblRendering.Visible = false;
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

            // Cancel any render still in flight. The previous CTS is disposed
            // only by its owning task after observing cancellation, so we just
            // signal here and replace the field.
            CancellationTokenSource previousCts = _renderCts;
            if (previousCts != null)
            {
                try { previousCts.Cancel(); } catch (ObjectDisposedException) { }
            }
            var myCts = new CancellationTokenSource();
            _renderCts = myCts;
            CancellationToken token = myCts.Token;

            // Start (or restart) the 500 ms indicator countdown. If this render
            // completes within 500 ms the label is never shown; if it overruns,
            // the tick handler reveals the label. Restarting on every dispatch
            // means continuous typing does not show the indicator at all.
            StartIndicatorCountdown();

            // GenerateHighLight builds its input/output paths from
            // Path.GetTempPath() + parameters.FileName. Compute the FileName
            // up front so we can best-effort delete the output file from the
            // cancellation path even though highlight.exe should not have
            // produced one (File.Delete is idempotent).
            string paneFileName = "preview-" + Guid.NewGuid().ToString("N");
            string expectedOutputPath = Path.Combine(Path.GetTempPath(), paneFileName) + ".html";

            Task.Run(() =>
            {
                string wrapped;
                string scratchPath = null;
                try
                {
                    string sessionDir = GetOrCreateSessionTempDir();
                    string scratchName = Path.GetRandomFileName();
                    scratchPath = Path.Combine(sessionDir, scratchName);

                    File.WriteAllText(scratchPath, snapshot.Content ?? string.Empty, Encoding.UTF8);

                    var paneParameters = new HighLightParameter
                    {
                        Content = snapshot.Content,
                        CodeType = snapshot.CodeType,
                        HighLightStyle = snapshot.HighLightStyle,
                        ShowLineNumber = snapshot.ShowLineNumber,
                        FileName = paneFileName,
                        Font = snapshot.Font,
                        FontSize = snapshot.FontSize,
                        HighlightColor = snapshot.HighlightColor
                    };

                    var generate = new GenerateHighLight(addinDir);
                    string outputPath = generate.GenerateHighLightCode(paneParameters, token);

                    string rawHtml = File.ReadAllText(outputPath, new UTF8Encoding(false));
                    try { File.Delete(outputPath); } catch { }
                    try { File.Delete(scratchPath); } catch { }

                    wrapped = PreviewHtmlWrapper.Wrap(rawHtml, darkMode, BackColor);
                }
                catch (OperationCanceledException)
                {
                    // Cancellation path: the child was killed mid-run, so the
                    // output file is unlikely to exist. Still best-effort delete
                    // both files since File.Delete on a missing path is a no-op
                    // after the existence check.
                    try { if (File.Exists(expectedOutputPath)) File.Delete(expectedOutputPath); } catch { }
                    try { if (scratchPath != null && File.Exists(scratchPath)) File.Delete(scratchPath); } catch { }
                    // Hide the indicator if no fresher render has started. If
                    // one has, the new render's StartIndicatorCountdown call
                    // already governs visibility - leave it alone.
                    _uiContext.Post(_ =>
                    {
                        if (IsDisposed) return;
                        if (mySeq == _renderSeq) HideIndicator();
                    }, null);
                    return;
                }
                catch
                {
                    // Phase 2.6 will surface errors into the pane. For Phase 1
                    // we silently leave the previous render on screen.
                    try { if (scratchPath != null && File.Exists(scratchPath)) File.Delete(scratchPath); } catch { }
                    _uiContext.Post(_ =>
                    {
                        if (IsDisposed) return;
                        if (mySeq == _renderSeq) HideIndicator();
                    }, null);
                    return;
                }
                finally
                {
                    // The CTS belongs to this render only; dispose once the
                    // task is done with it. If a newer render has already
                    // swapped _renderCts, that swap does not invalidate our
                    // local handle.
                    try { myCts.Dispose(); } catch { }
                }

                _uiContext.Post(_ =>
                {
                    if (mySeq != _renderSeq) return;
                    if (IsDisposed) return;
                    HideIndicator();
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

                if (_indicatorTimer != null)
                {
                    try { _indicatorTimer.Stop(); } catch { }
                    try { _indicatorTimer.Dispose(); } catch { }
                }

                // Cancel any in-flight render so the worker task does not
                // continue to drive a killed highlight.exe child after we
                // start tearing down the WebBrowser. The CTS itself is
                // disposed by the owning task in its finally block.
                CancellationTokenSource cts = _renderCts;
                if (cts != null)
                {
                    try { cts.Cancel(); } catch (ObjectDisposedException) { }
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
