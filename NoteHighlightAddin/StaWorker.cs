using System;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace NoteHighlightAddin
{
    /// <summary>
    /// Long-lived dedicated STA worker thread that owns a <see cref="BlockingCollection{T}"/> work
    /// queue, an <see cref="OneNoteMessageFilter"/> registration, and (transitively) the currently
    /// open WinForms dialog. All COM calls into OneNote from this add-in originate on this thread
    /// so that the marshalled <c>Application</c> proxy created on the worker's first cross-apartment
    /// call is reused for the worker's lifetime.
    ///
    /// See the threading-fix plan, Section 3 ("Architecture sketch") and Section 5 ("Step 2"),
    /// for the rationale behind the apartment-state, background-thread, exception-mode, queue
    /// disposal, and previous-filter-pointer ownership choices below.
    /// </summary>
    public sealed class StaWorker
    {
        private readonly string _name;
        private readonly OneNoteMessageFilter _filter;
        private readonly Action<Exception> _exceptionSink;
        private readonly BlockingCollection<Action> _queue = new BlockingCollection<Action>();

        private Thread _thread;
        private IntPtr _previousFilter = IntPtr.Zero;

        // Fatal HRESULTs that mean the host has gone away. On these we stop the worker without
        // surfacing a MessageBox: the host that would have rendered the box is no longer there
        // (and a MessageBox would either fail or hang waiting for an apartment that has died).
        private const int RPC_E_DISCONNECTED = unchecked((int)0x80010108);
        private const int RPC_E_SERVER_DIED = unchecked((int)0x80010012);
        private const int E_HANDLE = unchecked((int)0x80070006);

        /// <param name="name">
        /// Discriminator appended to the thread name for debugger diagnostics
        /// (e.g. "Main", "Settings"). Must not be null.
        /// </param>
        /// <param name="filter">
        /// The <see cref="OneNoteMessageFilter"/> instance to register on the worker STA. Must not
        /// be null. The worker calls <see cref="OneNoteMessageFilter.Register"/> after the apartment
        /// is established and stores the returned previous-filter <see cref="IntPtr"/> so it can be
        /// passed back to <see cref="OneNoteMessageFilter.Revoke"/> on shutdown.
        /// </param>
        /// <param name="exceptionSink">
        /// Optional callback invoked for non-fatal exceptions thrown by queued actions and for
        /// <see cref="Application.ThreadException"/>. If <c>null</c>, the default sink shows a
        /// <see cref="MessageBox"/>; this is safe because the worker is pumping its own message
        /// loop while an action runs.
        /// </param>
        public StaWorker(string name, OneNoteMessageFilter filter, Action<Exception> exceptionSink = null)
        {
            if (name == null) throw new ArgumentNullException("name");
            if (filter == null) throw new ArgumentNullException("filter");

            _name = name;
            _filter = filter;
            _exceptionSink = exceptionSink ?? DefaultExceptionSink;
        }

        /// <summary>
        /// True once <see cref="Start"/> has been called and the underlying thread is still alive.
        /// </summary>
        public bool IsRunning
        {
            get { return _thread != null && _thread.IsAlive; }
        }

        /// <summary>
        /// Creates and starts the underlying STA background thread. Must be called once and only
        /// once per instance.
        /// </summary>
        public void Start()
        {
            if (_thread != null)
            {
                throw new InvalidOperationException("StaWorker has already been started.");
            }

            _thread = new Thread(Run);

            // Apartment state and IsBackground MUST be set before Thread.Start.
            // - STA: required so the worker can host a WinForms message loop and own the marshalled
            //   OneNote Application proxy that originates on its first cross-apartment call.
            // - IsBackground = true: required for acceptance criterion A3. If Join later times out,
            //   a background worker thread does not keep ONENOTE.EXE alive. A foreground worker
            //   would cause exactly the ghost-process symptom A3 is trying to eliminate.
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.IsBackground = true;
            _thread.Name = "NoteHighlight STA Worker (" + _name + ")";

            _thread.Start();
        }

        /// <summary>
        /// Enqueues an action for execution on the worker. Swallows
        /// <see cref="InvalidOperationException"/> if the queue has been completed (i.e. the worker
        /// is shutting down) so callers do not need to guard against the shutdown race.
        /// </summary>
        public void Post(Action action)
        {
            if (action == null) throw new ArgumentNullException("action");
            try
            {
                _queue.Add(action);
            }
            catch (InvalidOperationException)
            {
                // CompleteAdding has been called - the worker is shutting down. Silently drop the
                // action; the caller (e.g. OnBeginShutdown / OnDisconnection) will not care.
            }
        }

        /// <summary>
        /// Marks the queue complete and waits up to <paramref name="joinTimeout"/> for the worker
        /// to drain its remaining actions and exit. Returns <c>true</c> iff the thread joined
        /// before the timeout. Caller is responsible for not calling <see cref="Post"/> after
        /// <see cref="Stop"/>.
        ///
        /// The internal <see cref="BlockingCollection{T}"/> is not disposed here: it is wrapped in
        /// a <c>using</c> on the worker thread itself, so its <see cref="IDisposable.Dispose"/>
        /// runs on the worker after the consuming loop exits. Disposing it from <see cref="Stop"/>
        /// would race with <c>GetConsumingEnumerable</c>.
        /// </summary>
        public bool Stop(TimeSpan joinTimeout)
        {
            if (_thread == null)
            {
                // Never started - trivially "joined".
                return true;
            }

            try
            {
                _queue.CompleteAdding();
            }
            catch (ObjectDisposedException)
            {
                // Already disposed by the worker. Nothing more to do.
            }

            return _thread.Join(joinTimeout);
        }

        private void Run()
        {
            // SetUnhandledExceptionMode MUST come before any Application.ThreadException
            // subscription, otherwise the documented exception-routing behaviour does not apply
            // for this thread.
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += OnThreadException;

            try
            {
                OneNoteMessageFilter.Register(_filter, out _previousFilter);

                using (_queue)
                {
                    foreach (var action in _queue.GetConsumingEnumerable())
                    {
                        try
                        {
                            action();
                        }
                        catch (COMException ex) when (IsFatalHostGone(ex.HResult))
                        {
                            // Host has disappeared. Stop the worker silently - a MessageBox here
                            // would either fail or hang. Break out of the consuming loop; the
                            // finally below will revoke the message filter.
                            SafeInvokeSink(ex);
                            break;
                        }
                        catch (Exception ex)
                        {
                            // Per-action try/catch keeps the worker alive across user-induced or
                            // transient exceptions. The next queued action still runs.
                            SafeInvokeSink(ex);
                        }
                    }
                }
            }
            finally
            {
                // Revoke and release the previous-filter pointer even on unexpected exit, so we
                // do not leak the previously-registered filter (which may belong to Office or
                // another add-in) for the life of the process.
                try
                {
                    OneNoteMessageFilter.Revoke(_previousFilter);
                }
                catch
                {
                    // Best-effort: never throw from the worker shutdown path.
                }
                _previousFilter = IntPtr.Zero;

                Application.ThreadException -= OnThreadException;
            }
        }

        private void OnThreadException(object sender, ThreadExceptionEventArgs e)
        {
            SafeInvokeSink(e.Exception);
        }

        private void SafeInvokeSink(Exception ex)
        {
            try
            {
                _exceptionSink(ex);
            }
            catch
            {
                // Swallow: an exception thrown from the sink itself must not tear down the worker.
            }
        }

        private static bool IsFatalHostGone(int hresult)
        {
            return hresult == RPC_E_DISCONNECTED
                || hresult == RPC_E_SERVER_DIED
                || hresult == E_HANDLE;
        }

        private static void DefaultExceptionSink(Exception ex)
        {
            // Safe to call MessageBox.Show from the worker: the worker is pumping its own message
            // loop while a queued action runs, so this dialog is parented to the worker STA and
            // does not require the main STA to pump on its behalf.
            MessageBox.Show("NoteHighlight worker error: " + ex.Message);
        }
    }
}
