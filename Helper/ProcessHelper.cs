using System;
using System.Text;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Threading;

namespace Helper
{
    public class ProcessHelper
    {
        #region -- Field and Property --

        public string WorkingDirectory { get; set; }

        public string FileName { get; set; }

        public string Arguments { get; set; }

        public bool IsWaitForInputIdle { get; set; }

        public ProcessWindowStyle WindowStyle { get; set; }

        /// <summary>
        /// Maximum time to wait for the child process to exit before it is forcibly
        /// terminated. Defaults to 30 seconds; callers may shorten or lengthen as
        /// required. A value of <see cref="Timeout.Infinite"/> disables the timeout
        /// (not recommended for the OneNote STA worker, which would otherwise hang).
        /// </summary>
        public int TimeoutMilliseconds { get; set; }

        /// <summary>
        /// Maximum number of stderr bytes (in UTF-16 chars) captured for inclusion
        /// in thrown exceptions. Output beyond this is silently discarded so that a
        /// runaway child cannot exhaust memory in the add-in process.
        /// </summary>
        private const int MaxCapturedStderrChars = 4096;

        /// <summary>
        /// 取得組件所在路徑
        /// Get Assembly Location
        /// </summary>
        public static string GetAssemblyLocationDirectory
        {
            get
            {
                string assemblyDirectory = Path.GetDirectoryName(Assembly.GetCallingAssembly().Location);
                return assemblyDirectory;
            }
        }


        #endregion

        #region -- Constructor --

        public ProcessHelper(string fileName):this(GetAssemblyLocationDirectory,fileName){}

        public ProcessHelper(string workingDirectory, string fileName)
        {
            WorkingDirectory = workingDirectory;
            FileName = fileName;
            TimeoutMilliseconds = 30000;
        }

        public ProcessHelper(string fileName, string[] arguments): this(GetAssemblyLocationDirectory, fileName, arguments){}

        public ProcessHelper(string workingDirectory, string fileName,string[] arguments)
        {
            WorkingDirectory = workingDirectory;
            FileName = fileName;
            Arguments = String.Join(" ", arguments);
            TimeoutMilliseconds = 30000;
        }

        #endregion

        /// <summary>
        /// Launches the configured executable and waits for it to terminate.
        ///
        /// Behaviour:
        ///  - <c>UseShellExecute</c> is set to <c>false</c> so stderr can be redirected
        ///    and no console window is spawned.
        ///  - stderr is captured asynchronously via <c>BeginErrorReadLine</c> to avoid
        ///    deadlocking on a full pipe buffer; up to ~4 KB is retained for diagnostics.
        ///  - The wait is bounded by <see cref="TimeoutMilliseconds"/>. On overrun the
        ///    process is killed and a <see cref="TimeoutException"/> is thrown.
        ///  - A non-zero <c>ExitCode</c> raises an <see cref="InvalidOperationException"/>
        ///    that surfaces the exit code and the first ~4 KB of captured stderr.
        ///  - <see cref="IsWaitForInputIdle"/> is still honoured; per its contract it
        ///    only applies to GUI processes and is a no-op for console children such as
        ///    <c>highlight.exe</c>.
        /// </summary>
        public void ProcessStart()
        {
            // Resolve FileName against WorkingDirectory when it is not already
            // an absolute path. With UseShellExecute = false (required for stderr
            // capture and CreateNoWindow), CreateProcess searches PATH only - it
            // does NOT consult ProcessStartInfo.WorkingDirectory when locating the
            // executable. A bare "highlight.exe" would therefore raise
            // Win32Exception 0x2 (ERROR_FILE_NOT_FOUND) even though the binary
            // sits beside the add-in in <addinDir>\highlight\.
            string resolvedFileName = FileName;
            if (!String.IsNullOrEmpty(resolvedFileName)
                && !Path.IsPathRooted(resolvedFileName)
                && !String.IsNullOrEmpty(WorkingDirectory))
            {
                resolvedFileName = Path.Combine(WorkingDirectory, resolvedFileName);
            }

            if (!String.IsNullOrEmpty(resolvedFileName) && !File.Exists(resolvedFileName))
            {
                throw new FileNotFoundException(
                    "Executable not found at resolved path '" + resolvedFileName +
                    "' (WorkingDirectory='" + WorkingDirectory + "', FileName='" +
                    FileName + "').", resolvedFileName);
            }

            using (Process p = new Process())
            {
                ProcessStartInfo info = new ProcessStartInfo();
                info.WorkingDirectory = WorkingDirectory;
                info.FileName = resolvedFileName;
                info.Arguments = Arguments;
                info.WindowStyle = WindowStyle;
                info.UseShellExecute = false;
                info.CreateNoWindow = true;
                info.RedirectStandardError = true;
                // Note: stdout is left attached to the (suppressed) console; highlight.exe
                // writes its result to the configured output file, so redirecting stdout
                // would only add a pipe to drain. If a future caller relies on stdout, add
                // a parallel BeginOutputReadLine handler here.

                p.StartInfo = info;

                StringBuilder stderrBuffer = new StringBuilder();
                object stderrLock = new object();
                p.ErrorDataReceived += (sender, e) =>
                {
                    if (e.Data == null) return;
                    lock (stderrLock)
                    {
                        int remaining = MaxCapturedStderrChars - stderrBuffer.Length;
                        if (remaining <= 0) return;
                        // +1 for newline separator between lines.
                        if (e.Data.Length + 1 <= remaining)
                        {
                            stderrBuffer.Append(e.Data).Append('\n');
                        }
                        else
                        {
                            // Truncate cleanly at the buffer cap rather than overflowing.
                            int take = Math.Max(0, remaining - 1);
                            if (take > 0) stderrBuffer.Append(e.Data, 0, take);
                            stderrBuffer.Append('\n');
                        }
                    }
                };

                p.Start();
                p.BeginErrorReadLine();

                if (IsWaitForInputIdle)
                {
                    // WaitForInputIdle is only valid for GUI processes; guard so a
                    // console child does not raise InvalidOperationException.
                    try { p.WaitForInputIdle(); } catch (InvalidOperationException) { }
                }

                int timeout = TimeoutMilliseconds <= 0 ? 30000 : TimeoutMilliseconds;
                bool exited = p.WaitForExit(timeout);
                if (!exited)
                {
                    try { p.Kill(); } catch { /* already gone or access denied */ }
                    // Give the async stderr reader a brief window to drain whatever
                    // was buffered before the kill so the thrown message is useful.
                    try { p.WaitForExit(500); } catch { }

                    string stderrSnapshot;
                    lock (stderrLock) { stderrSnapshot = stderrBuffer.ToString(); }
                    throw new TimeoutException(
                        String.Format(
                            "Process '{0}' did not exit within {1} ms and was terminated. Stderr: {2}",
                            FileName, timeout,
                            String.IsNullOrEmpty(stderrSnapshot) ? "(none)" : stderrSnapshot));
                }

                // Ensure all asynchronous stderr events have been raised before we
                // read the buffer (WaitForExit() with no argument flushes them).
                p.WaitForExit();

                if (p.ExitCode != 0)
                {
                    string stderrSnapshot;
                    lock (stderrLock) { stderrSnapshot = stderrBuffer.ToString(); }
                    throw new InvalidOperationException(
                        String.Format(
                            "Process '{0}' exited with code {1}. Stderr: {2}",
                            FileName, p.ExitCode,
                            String.IsNullOrEmpty(stderrSnapshot) ? "(none)" : stderrSnapshot));
                }
            }
        }


        public static string GetDirectoryFromPath(string path)
        {
            string assemblyDirectory = Path.GetDirectoryName(path);
            return assemblyDirectory;

        }
    }
}
