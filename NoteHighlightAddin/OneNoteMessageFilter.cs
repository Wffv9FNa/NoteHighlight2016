using System;
using System.Runtime.InteropServices;

namespace NoteHighlightAddin
{
    /// <summary>
    /// Managed declaration of the COM <c>IOleMessageFilter</c> (a.k.a. <c>IMessageFilter</c>) interface
    /// used to handle COM call rejections / retry-later responses from a busy Office STA host.
    ///
    /// The IID 00000016-0000-0000-C000-000000000046 names the same COM v-table that Microsoft samples
    /// refer to as both "IMessageFilter" and "IOleMessageFilter"; the <c>IOle*</c> name is used here to
    /// disambiguate from <see cref="System.Windows.Forms.IMessageFilter"/>.
    /// </summary>
    [ComImport]
    [Guid("00000016-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(
            int dwCallType,
            IntPtr hTaskCaller,
            int dwTickCount,
            IntPtr lpInterfaceInfo);

        [PreserveSig]
        int RetryRejectedCall(
            IntPtr hTaskCallee,
            int dwTickCount,
            int dwRejectType);

        [PreserveSig]
        int MessagePending(
            IntPtr hTaskCallee,
            int dwTickCount,
            int dwPendingType);
    }

    /// <summary>
    /// Managed implementation of <see cref="IOleMessageFilter"/> that retries OneNote COM calls with
    /// exponential backoff when OneNote's main STA is busy (e.g. mid-sync, modal dialog, etc.).
    ///
    /// The filter is registered on the dedicated STA worker thread that originates the COM calls,
    /// never on OneNote's own main STA. See the threading-fix plan, section 3, for the design.
    /// </summary>
    public sealed class OneNoteMessageFilter : IOleMessageFilter
    {
        // dwRejectType values
        private const int SERVERCALL_ISHANDLED = 0;
        private const int SERVERCALL_REJECTED = 0;
        private const int SERVERCALL_RETRYLATER = 2;

        // dwPendingType / MessagePending return values
        private const int PENDINGMSG_WAITDEFPROCESS = 2;

        // Per-retry backoff cap (per the plan, each individual retry is capped at 2000 ms).
        private const int PerRetryCapMs = 2000;

        // Tracks cumulative wait emitted in the current retry sequence. Reset whenever the filter
        // returns -1 (the call is being cancelled, so the next rejection observed must belong to a
        // fresh outbound call). Tracking by hTaskCallee/dwTickCount would be more rigorous but is
        // unnecessary here: the filter is registered on a single dedicated STA worker that issues
        // strictly serial outbound calls (one Application.Run nested loop / queued action at a time).
        private int _cumulativeWaitMs;

        // The next backoff to emit on a busy/rejected response.
        private int _nextBackoffMs;

        // Set by SignalShutdown so an in-flight retry loop bails out immediately rather than
        // extending OneNote shutdown by up to MaxRetryMilliseconds.
        private volatile bool _shuttingDown;

        /// <summary>
        /// Total cumulative wait budget across all retries of a single outbound call, in
        /// milliseconds. Public and mutable so the manual negative-test in Step 4 can set it to 0
        /// at runtime to force the filter to cancel immediately.
        /// </summary>
        public int MaxRetryMilliseconds { get; set; } = 30000;

        /// <summary>
        /// Initial backoff in milliseconds. Doubled on each subsequent retry, capped per-retry at
        /// 2000 ms. Public and mutable for the same reason as <see cref="MaxRetryMilliseconds"/>.
        /// </summary>
        public int BackoffStepMilliseconds { get; set; } = 100;

        public OneNoteMessageFilter()
        {
            ResetSequence();
        }

        /// <summary>
        /// Signals the filter to stop retrying. Called from <c>OnBeginShutdown</c> / <c>OnDisconnection</c>
        /// on the main STA. Uses a volatile flag so the worker STA observes the write without locking.
        /// </summary>
        public void SignalShutdown()
        {
            _shuttingDown = true;
        }

        public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            // The worker STA exposes no COM interfaces, so this is never actually invoked. Returning
            // SERVERCALL_ISHANDLED is the documented safe choice.
            return SERVERCALL_ISHANDLED;
        }

        public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            if (_shuttingDown)
            {
                ResetSequence();
                return -1;
            }

            // SERVERCALL_REJECTED (0) and SERVERCALL_RETRYLATER (2) are the only documented values
            // we should retry on. SERVERCALL_REJECTED and SERVERCALL_ISHANDLED share the value 0;
            // SERVERCALL_ISHANDLED is never passed to RetryRejectedCall by COM, so treating 0 as
            // "rejected" here is correct in this direction.
            if (dwRejectType != SERVERCALL_RETRYLATER && dwRejectType != SERVERCALL_REJECTED)
            {
                ResetSequence();
                return -1;
            }

            int wait = _nextBackoffMs;
            if (wait > PerRetryCapMs)
            {
                wait = PerRetryCapMs;
            }

            // If adding this wait would exceed the configured budget, give up.
            if (_cumulativeWaitMs + wait > MaxRetryMilliseconds)
            {
                ResetSequence();
                return -1;
            }

            _cumulativeWaitMs += wait;

            // Prepare next step: double, but cap so we do not overflow and so the next read of
            // _nextBackoffMs is also clamped to the per-retry cap.
            int doubled = _nextBackoffMs * 2;
            if (doubled < _nextBackoffMs || doubled > PerRetryCapMs)
            {
                doubled = PerRetryCapMs;
            }
            _nextBackoffMs = doubled;

            return wait;
        }

        public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            return PENDINGMSG_WAITDEFPROCESS;
        }

        private void ResetSequence()
        {
            _cumulativeWaitMs = 0;
            int step = BackoffStepMilliseconds;
            if (step < 1)
            {
                step = 1;
            }
            _nextBackoffMs = step;
        }

        // P/Invoke. The previous-filter out-parameter is declared as IntPtr (not IOleMessageFilter)
        // because COM has already AddRef'd the returned object; we must release it explicitly via
        // Marshal.Release on shutdown to avoid leaking the previously-registered filter (which may
        // belong to Office itself or another add-in) for the life of the process.
        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(
            IOleMessageFilter newFilter,
            out IntPtr ppPrevFilter);

        /// <summary>
        /// Installs <paramref name="filter"/> as the message filter for the current STA and returns
        /// the previously-registered filter pointer (which must be passed to <see cref="Revoke"/> on
        /// shutdown so it can be released).
        /// </summary>
        public static int Register(OneNoteMessageFilter filter, out IntPtr previous)
        {
            if (filter == null) throw new ArgumentNullException("filter");
            return CoRegisterMessageFilter(filter, out previous);
        }

        /// <summary>
        /// Unregisters whatever filter is currently installed on this STA, then releases the
        /// previously-saved pointer if non-zero. Best-effort: callers should not throw on failure.
        /// </summary>
        public static void Revoke(IntPtr previous)
        {
            IntPtr ignored;
            CoRegisterMessageFilter(null, out ignored);
            if (previous != IntPtr.Zero)
            {
                Marshal.Release(previous);
            }
        }
    }
}
