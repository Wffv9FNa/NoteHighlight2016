/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Extensibility;
using Microsoft.Office.Core;
using NoteHighlightAddin.Utilities;
using Application = Microsoft.Office.Interop.OneNote.Application;  // Conflicts with System.Windows.Forms
using System.Reflection;
using System.Drawing;
using Microsoft.Office.Interop.OneNote;
using NoteHighlightForm;
using System.Text;
using System.Linq;
using Helper;
using System.Threading;
using System.Web;
using GenerateHighlightContent;
using System.Configuration;
using System.Globalization;
using NoteHighlightAddin.Preview;

#pragma warning disable CS3003 // Type is not CLS-compliant

namespace NoteHighlightAddin
{
	[ComVisible(true)]
	[Guid("4C6B0362-F139-417F-9661-3663C268B9E9"), ProgId("NoteHighlight2016.AddIn")]
	public class AddIn : IDTExtensibility2, IRibbonExtensibility
	{
		protected Application OneNoteApplication
		{ get; set; }

        public XNamespace ns;

        // Worker infrastructure. Ribbon callbacks are serialised by Office onto the main UI thread,
        // so plain null-checks are sufficient for lazy construction - no lock, no Lazy<T>, no
        // double-checked locking. See plan section 3 ("Ownership and lifecycle").
        private StaWorker _mainWorker;
        private StaWorker _settingsWorker;

        // The "currently open" form pointers. These fields are written and read ONLY on the
        // corresponding worker thread; the main STA never touches them. As a result no volatile /
        // Interlocked is required. Cross-apartment close is achieved by posting an action onto the
        // worker's own queue (see OnBeginShutdown).
        private MainForm _currentMainForm;
        private SettingsForm _currentSettingsForm;
        private LanguageSettingsForm _currentLanguagesForm;

        // Filters are held by the AddIn so SignalShutdown() can be invoked from OnBeginShutdown /
        // OnDisconnection on the main STA. The filter's _shuttingDown flag is volatile, so this
        // cross-thread write is safe without additional synchronisation.
        private OneNoteMessageFilter _mainFilter;
        private OneNoteMessageFilter _settingsFilter;

        private bool QuickStyle { get; set; }

        private bool DarkMode { get; set; }

        // Idempotence guard for the Phase 2.4 stale-temp-dir sweep. OnConnection
        // may legitimately fire more than once over the add-in lifetime (e.g. a
        // host reconnect); the sweep is best-effort and only needs to run once
        // per process. Plain int + Interlocked.CompareExchange avoids the lock
        // and is sufficient here - OnConnection is always called on the main
        // STA, but the sweep itself dispatches to the thread pool.
        private static int _sweepDone;

        // Guard for AssemblyResolve subscription. OnConnection is normally called
        // once per AppDomain lifetime, but a defensive flag keeps us from stacking
        // handlers if the host ever reconnects. Plain bool is fine: OnConnection
        // runs on the main STA and the flag is only ever set there.
        private static bool _assemblyResolveSubscribed;

        // Language-picker state (Phase 1). _ribbon is captured by OnRibbonLoad and
        // released in OnDisconnection. _languages is the per-user picker state;
        // reads/writes are coordinated by _languagesLock. _invalidatePending is
        // set from any thread when settings change; only an onAction observer
        // ever calls _ribbon.Invalidate() - never a getVisible / getContent
        // callback (re-entrant Invalidate from inside a get* callback is at best
        // a no-op and at worst stalls the ribbon for the rest of the session).
        //
        // Threading reality (Phase 0 spike, 2026-05-13): OneNote routes ribbon
        // callbacks through MTA threads, not a single STA. _ribbon is agile from
        // MTA, so the cached RCW is callable from any ribbon-callback thread.
        // SynchronizationContext.Current is null on every ribbon-callback thread
        // observed, so capturing it for .Post(...) would NRE - we use the flag
        // instead.
        private Microsoft.Office.Core.IRibbonUI _ribbon;
        private LanguageSettings _languages;
        private readonly object _languagesLock = new object();
        private volatile bool _invalidatePending;

        public AddIn()
		{
		}

        /// <summary>
        /// Returns the directory containing NoteHighlightAddin.dll. Prefer
        /// Assembly.Location, but fall back to CodeBase for hosts that load the
        /// assembly from a byte array (Location returns string.Empty in that
        /// case). This is the authoritative install path used by callers that
        /// need to locate sibling files (highlight.exe, themes, dll.config).
        /// </summary>
        internal static string GetAddinDirectory()
        {
            var asm = typeof(AddIn).Assembly;
            string loc = null;
            try
            {
                loc = asm.Location;
            }
            catch (NotSupportedException) { /* dynamic assembly */ }

            if (string.IsNullOrEmpty(loc))
            {
                try
                {
                    var cb = asm.CodeBase;
                    if (!string.IsNullOrEmpty(cb))
                        loc = new Uri(cb).LocalPath;
                }
                catch (NotSupportedException) { /* dynamic assembly */ }
                catch (UriFormatException) { /* malformed CodeBase URI */ }
                catch (Exception)
                {
                    // Defensive: never let path-discovery throw out of a ribbon callback.
                }
            }

            return string.IsNullOrEmpty(loc) ? null : Path.GetDirectoryName(loc);
        }

        /// <summary>
        /// AssemblyResolve handler that locates GenerateHighlightContent.dll next
        /// to the add-in DLL. Required because, under OneNote COM activation, the
        /// AppDomain's AppBase is OneNote's Office16 directory rather than the
        /// add-in's bin folder, so Fusion cannot find sibling DLLs via its normal
        /// partial-name probe. Returning the loaded assembly from this handler
        /// grafts it onto the Load context, which is what ConfigurationManager's
        /// partial-name Type.GetType lookup consults. Eager Assembly.LoadFrom in
        /// OnConnection does NOT achieve the same thing - LoadFrom-context
        /// assemblies are invisible to the Load-context partial-name probe.
        /// MUST NEVER throw: an exception out of AssemblyResolve tears down the
        /// AppDomain. Always log + return null on failure.
        /// </summary>
        private static Assembly ResolveGenerateHighlightContent(object sender, ResolveEventArgs args)
        {
            try
            {
                if (args == null || string.IsNullOrEmpty(args.Name))
                    return null;

                // Match both the partial name ("GenerateHighlightContent") and any
                // strong-name variant ("GenerateHighlightContent, Version=..."). Use
                // Ordinal because this handler fires for every assembly load.
                if (!args.Name.StartsWith("GenerateHighlightContent", StringComparison.Ordinal))
                    return null;

                var addinDir = GetAddinDirectory();
                if (string.IsNullOrEmpty(addinDir))
                {
                    System.Diagnostics.Trace.TraceError("NoteHighlight2016: AssemblyResolve for '" + args.Name + "' failed - GetAddinDirectory() returned null.");
                    return null;
                }

                var ghcPath = Path.Combine(addinDir, "GenerateHighlightContent.dll");
                if (!File.Exists(ghcPath))
                {
                    System.Diagnostics.Trace.TraceError("NoteHighlight2016: AssemblyResolve for '" + args.Name + "' failed - file not found at resolved path '" + ghcPath + "'.");
                    return null;
                }

                var loaded = Assembly.LoadFrom(ghcPath);
                System.Diagnostics.Trace.TraceInformation("NoteHighlight2016: AssemblyResolve loaded '" + args.Name + "' from '" + ghcPath + "' (FullName=" + loaded.FullName + ").");
                return loaded;
            }
            catch (Exception ex)
            {
                // Defensive catch-all: never let an exception escape AssemblyResolve.
                string attempted;
                try
                {
                    var addinDir = GetAddinDirectory();
                    attempted = string.IsNullOrEmpty(addinDir) ? "<null>" : Path.Combine(addinDir, "GenerateHighlightContent.dll");
                }
                catch
                {
                    attempted = "<path-resolution-failed>";
                }
                System.Diagnostics.Trace.TraceError("NoteHighlight2016: AssemblyResolve for '" + (args != null ? args.Name : "<null>") + "' threw (attempted path '" + attempted + "'). " + ex.ToString());
                return null;
            }
        }

		/// <summary>
		/// Returns the XML in Ribbon.xml so OneNote knows how to render our ribbon
		/// </summary>
		/// <param name="RibbonID"></param>
		/// <returns></returns>
		public string GetCustomUI(string RibbonID)
		{
            return LoadRibbon();

        }

        private string LoadRibbon()
        {
            // ribbon.xml is embedded into the assembly (via Properties\Resources.resx
            // -> Resources.ribbon) because it is the unversioned-file class of MSI
            // upgrade bug: Windows Installer skips replacing unversioned files when
            // their mtime != ctime on disk, so an MSI upgrade would silently leave
            // an older ribbon.xml in place. The DLL is versioned and always replaced
            // on upgrade, so the embedded copy moves with it. See
            // .local/docs/bugs/msi-ribbon-stale-on-upgrade.md (bug 13.1).
            //
            // Fallback order (post bug 13.1 followup, 2026-05-13 evening):
            //   1. On-disk ribbon.override.xml next to the DLL, if present.
            //      Dev / diagnostic hot-edit convenience: edit the file, restart
            //      OneNote, no rebuild. End-user MSIs never ship this filename.
            //   1b. A legacy on-disk ribbon.xml (from a pre-3.9 install whose
            //       Permanent=TRUE MSI component never uninstalled the file) is
            //       deliberately IGNORED with a Trace warning. It is almost
            //       certainly stale; reading it would beat the embedded copy
            //       and silently regress ribbon changes shipped via DLL upgrade.
            //   2. Embedded Resources.ribbon (the canonical end-user path).
            //   3. Empty string + MessageBox (existing error path).

            try
            {
                var addinDir = GetAddinDirectory();
                if (!string.IsNullOrEmpty(addinDir))
                {
                    // Legacy-file safety net (1b). Log the resolved path so a user
                    // who wants a clean install knows where to point a Remove-Item.
                    var legacy = Path.Combine(addinDir, LanguageRegistry.LegacyOnDiskRibbonFileName);
                    if (File.Exists(legacy))
                    {
                        System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: ignoring legacy on-disk '" + legacy + "' - the embedded ribbon resource is canonical. You may delete this file manually; the add-in no longer reads it.");
                    }

                    var onDisk = Path.Combine(addinDir, LanguageRegistry.OnDiskRibbonFileName);
                    // File.Exists guard mirrors the rule in feedback_com_addin_path_traps.md -
                    // never assume a sibling file is present under COM activation.
                    if (File.Exists(onDisk))
                    {
                        System.Diagnostics.Trace.TraceInformation("NoteHighlight2016: using on-disk ribbon override '" + onDisk + "'.");
                        return File.ReadAllText(onDisk);
                    }
                }
            }
            catch (Exception e)
            {
                // Hot-edit fallback failure is non-fatal; fall through to the embedded copy.
                System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: on-disk ribbon override read failed; falling back to embedded resource. " + e.Message);
            }

            try
            {
                var embedded = Properties.Resources.ribbon;
                if (!string.IsNullOrEmpty(embedded))
                    return embedded;

                MessageBox.Show("Exception from Addin.LoadRibbon: embedded ribbon resource was empty.");
                return "";
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception from Addin.LoadRibbon:" + e.Message);
                return "";
            }
        }

        public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>
		/// Cleanup. Must return promptly - OneNote expects OnBeginShutdown to be non-blocking.
		/// Close is delivered via the form's own BeginInvoke (asynchronous, fire-and-forget) so it
		/// lands on the WinForms message pump that Application.Run is actively pumping. Posting to
		/// the worker's BlockingCollection here would never fire while the worker is parked inside
		/// Application.Run, since the queue is only drained by the outer foreach in StaWorker.Run.
		/// BeginInvoke (unlike synchronous Invoke) does not block the main STA, so the
		/// "main-STA-waits-on-worker-mid-COM-call" deadlock does not apply.
		/// </summary>
		/// <param name="custom"></param>
		public void OnBeginShutdown(ref Array custom)
		{
			_mainFilter?.SignalShutdown();
			_settingsFilter?.SignalShutdown();

			var mf = _currentMainForm;
			if (mf != null)
			{
				try { mf.BeginInvoke(new Action(() => { try { mf.Close(); } catch { } })); }
				catch { /* form may already be disposed or its handle not yet created */ }
			}

			var sf = _currentSettingsForm;
			if (sf != null)
			{
				try { sf.BeginInvoke(new Action(() => { try { sf.Close(); } catch { } })); }
				catch { /* form may already be disposed or its handle not yet created */ }
			}

			var lf = _currentLanguagesForm;
			if (lf != null)
			{
				try { lf.BeginInvoke(new Action(() => { try { lf.Close(); } catch { } })); }
				catch { /* form may already be disposed or its handle not yet created */ }
			}
			// Return immediately. The actual worker join happens in OnDisconnection.
		}

		/// <summary>
		/// Called upon startup.
		/// Keeps a reference to the current OneNote application object.
		/// </summary>
		/// <param name="application"></param>
		/// <param name="connectMode"></param>
		/// <param name="addInInst"></param>
		/// <param name="custom"></param>
		public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
		{
			// Plan section 4.2 / reviewer 3.9: ext_cm_UISetup is fired by Office to
			// let the add-in register its UI; no live application is attached and
			// per-user state must not be initialised. Bail before SetOneNoteApplication.
			if (ConnectMode == ext_ConnectMode.ext_cm_UISetup)
				return;

			// MUST subscribe before the first ConfigurationManager.GetSection("HighLightSection") call (MainForm.LoadThemes, SettingsForm, PreviewPane, etc). Under COM activation the AppDomain's AppBase is OneNote's Office16 directory, not the add-in's bin folder, so Fusion's partial-name probe (Type.GetType -> Assembly.Load) cannot locate GenerateHighlightContent.dll on its own. Eager Assembly.LoadFrom does NOT fix this because LoadFrom-context assemblies are invisible to the Load-context partial-name lookup; AssemblyResolve is the canonical workaround because the assembly it returns is treated as if Fusion had resolved it itself. Do NOT move config-section access into a type initialiser that runs before OnConnection - that re-introduces the partial-binding regression. See .local/plans/fix-broken-roundtrip-tests.md Option E.
			try
			{
				if (!_assemblyResolveSubscribed)
				{
					AppDomain.CurrentDomain.AssemblyResolve += ResolveGenerateHighlightContent;
					_assemblyResolveSubscribed = true;
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Trace.TraceError("NoteHighlight2016: failed to subscribe AssemblyResolve handler; ConfigurationManager.GetSection(\"HighLightSection\") will likely fail under COM activation. " + ex);
				// Never rethrow - COM activation will silently disable the add-in.
			}

			SetOneNoteApplication((Application)Application);

			// Language-picker bootstrap (Phase 1). Wrapped in defensive try/catch -
			// neither call must escape into OnConnection or Office aborts add-in load.
			try
			{
				// No-arg overload: prefers an on-disk ribbon.xml for dev hot-edit,
				// falls back to the embedded canonical copy. See bug 13.1 / Option B.
				LanguageRegistry.Initialise();
			}
			catch (Exception ex)
			{
				System.Diagnostics.Trace.TraceError("NoteHighlight2016: LanguageRegistry.Initialise threw; continuing with empty registry. " + ex);
			}

			try
			{
				var knownTags = LanguageRegistry.All.Select(d => d.Tag).ToList();
				var seeded = LanguageSettings.LoadOrSeed(
					SettingsHelper.LanguagesJsonPath,
					knownTags,
					LanguageRegistry.DefaultPinned);
				lock (_languagesLock)
				{
					_languages = seeded;
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Trace.TraceError("NoteHighlight2016: LanguageSettings.LoadOrSeed threw; ribbon will use fallback defaults. " + ex);
				// Fall back to an in-memory defaults instance so the slot get* callbacks
				// still have something to consult.
				try
				{
					var fallback = LanguageSettings.LoadOrSeed(null, null, LanguageRegistry.DefaultPinned);
					lock (_languagesLock)
					{
						_languages = fallback;
					}
				}
				catch
				{
					// Never throw out of OnConnection.
				}
			}

			// Phase 2.4: best-effort cleanup of stale preview-pane temp dirs left
			// behind by previous OneNote crashes / kills. Wrap the dispatch (not
			// just the sweep itself) so absolutely nothing can escape into
			// OnConnection - a throw out of this method aborts add-in load.
			try
			{
				if (Interlocked.CompareExchange(ref _sweepDone, 1, 0) == 0)
				{
					ThreadPool.QueueUserWorkItem(_ =>
					{
						try { SweepStalePreviewTempDirs(TimeSpan.FromHours(24)); }
						catch
						{
							// Best-effort. Never surface to the host - the worker
							// thread has no UI context and an unhandled exception
							// here would tear down the process under legacy
							// CLR policy.
						}
					});
				}
			}
			catch
			{
				// Defence in depth: ThreadPool.QueueUserWorkItem itself should
				// not throw, but if it ever does we still must not break the
				// add-in load path.
			}
		}

		/// <summary>
		/// Best-effort sweep of leftover per-session preview temp directories
		/// created by <see cref="Preview.PreviewPane"/>. Each pane allocates a
		/// directory named <c>{TempPath}\NoteHighlight2016\preview-{guid}</c>
		/// and removes it on dispose; OneNote crashes leave them behind. We
		/// delete only the <c>preview-*</c> children whose last-write time is
		/// older than <paramref name="age"/>, never the parent
		/// <c>NoteHighlight2016</c> directory itself (a live pane in another
		/// OneNote instance may have just created one).
		/// </summary>
		/// <remarks>
		/// Safe to call from any thread. All path resolution uses
		/// <see cref="Path.GetTempPath"/> - it does not touch
		/// <c>Assembly.GetExecutingAssembly().Location</c>, which is unsafe
		/// under COM activation (see feedback_com_addin_path_traps.md).
		/// Every per-directory operation is wrapped in try/catch so a single
		/// locked / permission-denied entry cannot abort the sweep.
		/// </remarks>
		private static void SweepStalePreviewTempDirs(TimeSpan age)
		{
			string root;
			try
			{
				root = Path.Combine(Path.GetTempPath(), "NoteHighlight2016");
			}
			catch
			{
				return;
			}

			if (string.IsNullOrEmpty(root)) return;
			if (!Directory.Exists(root)) return;

			string[] candidates;
			try
			{
				candidates = Directory.GetDirectories(root, "preview-*", SearchOption.TopDirectoryOnly);
			}
			catch
			{
				return;
			}

			DateTime cutoffUtc = DateTime.UtcNow - age;

			foreach (string dir in candidates)
			{
				try
				{
					// LastWriteTimeUtc is the most robust signal on Windows tmp
					// dirs - CreationTimeUtc is preserved across copies/moves
					// and can outlive the originating session. A live pane
					// touches its files continuously so LastWriteTime tracks
					// "in use" correctly.
					DateTime lastWriteUtc = Directory.GetLastWriteTimeUtc(dir);
					if (lastWriteUtc > cutoffUtc) continue;

					Directory.Delete(dir, true);
				}
				catch
				{
					// Best-effort: swallow IOException, UnauthorizedAccessException,
					// DirectoryNotFoundException (race), etc. Move on.
				}
			}
		}

		public void SetOneNoteApplication(Application application)
		{
			OneNoteApplication = application;
		}

		/// <summary>
		/// Cleanup
		/// </summary>
		/// <param name="RemoveMode"></param>
		/// <param name="custom"></param>
		public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
		{
			// 1. Signal both filters to stop retrying. Idempotent if OnBeginShutdown already ran.
			_mainFilter?.SignalShutdown();
			_settingsFilter?.SignalShutdown();

			// 2. Stop each worker (CompleteAdding + Join with a 5 s timeout). The returned bool
			//    tells us whether the worker exited cleanly; we use it to decide whether it is
			//    safe to release the RCW it might still be holding.
			bool mainJoined = true;
			bool settingsJoined = true;

			if (_mainWorker != null)
			{
				mainJoined = _mainWorker.Stop(TimeSpan.FromSeconds(5));
			}
			if (_settingsWorker != null)
			{
				settingsJoined = _settingsWorker.Stop(TimeSpan.FromSeconds(5));
			}

			// 3. Release the OneNote RCW only if the main worker joined cleanly. If it timed out,
			//    the worker may still be inside a COM call holding that proxy; releasing here would
			//    surface as InvalidComObjectException on the worker. The OS reclaims the RCW on
			//    process exit; leaving it unreleased is the safer choice during a stuck shutdown.
			//    ReleaseComObject (single decrement) not FinalReleaseComObject, per the plan.
			if (mainJoined && OneNoteApplication != null && Marshal.IsComObject(OneNoteApplication))
			{
				try
				{
					Marshal.ReleaseComObject(OneNoteApplication);
				}
				catch
				{
					// Best-effort: never throw out of OnDisconnection.
				}
			}

			// 4. Field-nulling MUST come after the join, not before. The worker may legitimately
			//    touch the proxy until the moment it exits.
			OneNoteApplication = null;

			// Release the cached IRibbonUI RCW. Mirrors the OneNoteApplication release above
			// (plan section 4.2 / reviewer 3.1). Must come after the worker join so a still-
			// running worker cannot observe a torn-down _ribbon. Best-effort: never throw.
			if (_ribbon != null)
			{
				try
				{
					if (Marshal.IsComObject(_ribbon))
						Marshal.ReleaseComObject(_ribbon);
				}
				catch (Exception e)
				{
					System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: ReleaseComObject(_ribbon): " + e);
				}
				_ribbon = null;
			}

			// (Suppress unused-variable warning for settingsJoined - it is intentionally captured
			// for future symmetry / diagnostics even though we have no settings RCW to release.)
			GC.KeepAlive(settingsJoined);

			// GC.Collect / GC.WaitForPendingFinalizers removed deliberately. They served no real
			// purpose here and could mask leaks (the cleaner ReleaseComObject above is what
			// actually matters).
		}

		public void OnStartupComplete(ref Array custom)
		{
		}

        /// <summary>
        /// <c>onLoad</c> handler declared on the <c>&lt;customUI&gt;</c> root. Office calls this
        /// once after parsing the ribbon XML; we cache the <see cref="IRibbonUI"/> RCW so later
        /// code can request a re-collection of <c>get*</c> values via <see cref="IRibbonUI.Invalidate"/>.
        ///
        /// <para>
        /// Threading (Phase 0 spike): this callback fires on an MTA thread, not an STA. The
        /// captured RCW is agile-from-MTA so subsequent <c>onAction</c> callbacks - which may
        /// run on different MTA threads - can use the same reference without RPC errors.
        /// </para>
        /// </summary>
        [System.CLSCompliant(false)]
        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }

        /// <summary>
        /// Parse a slot button id of the form <c>slotLangNN</c> into its integer index.
        /// Returns -1 on any malformed id (null, wrong prefix, non-numeric suffix).
        /// </summary>
        internal static int SlotIndexFromControlId(string controlId)
        {
            const string prefix = "slotLang";
            if (string.IsNullOrEmpty(controlId)) return -1;
            if (!controlId.StartsWith(prefix, StringComparison.Ordinal)) return -1;
            string suffix = controlId.Substring(prefix.Length);
            if (suffix.Length == 0) return -1;
            if (!int.TryParse(suffix, NumberStyles.None, CultureInfo.InvariantCulture, out int idx))
                return -1;
            return idx;
        }

        /// <summary>
        /// Resolve a slot index to the pinned <see cref="LanguageDescriptor"/>, or <c>null</c> if
        /// the slot is beyond the current pinned count or the pinned tag is not in the registry.
        ///
        /// <para>
        /// Thread-safety invariant (reviewer m1): reading <c>snap.Pinned</c> lock-free after the
        /// snapshot is only safe because the live <see cref="LanguageSettings"/> instance handed
        /// to the AddIn is never mutated in place - <see cref="LanguageSettingsForm"/> edits a
        /// clone and <see cref="ReplaceLanguageSettings"/> swaps the whole <c>_languages</c>
        /// reference under <see cref="_languagesLock"/>. <see cref="LanguageSettings"/> does
        /// expose mutators (Pin/Unpin/Reorder), so this invariant is load-bearing: a future change
        /// to the form's save path that mutated the live instance would quietly break it.
        /// </para>
        /// </summary>
        private LanguageDescriptor ResolveSlot(int slotIndex)
        {
            LanguageSettings snap;
            lock (_languagesLock) { snap = _languages; }
            return ResolveSlot(snap, LanguageRegistry.All, slotIndex);
        }

        /// <summary>
        /// Pure / static core of <see cref="ResolveSlot(int)"/>. Exposed so unit tests can exercise
        /// the slot-index-to-descriptor mapping with a hand-built descriptor list, without standing
        /// up the static <see cref="LanguageRegistry"/> (mirrors how <see cref="BuildMoreLanguagesMenuXml"/>
        /// was made static / testable). Returns <c>null</c> when <paramref name="settings"/> is null,
        /// the index is out of range, or the pinned tag is not present in <paramref name="registry"/>.
        /// </summary>
        internal static LanguageDescriptor ResolveSlot(LanguageSettings settings,
                                                       IReadOnlyList<LanguageDescriptor> registry,
                                                       int slotIndex)
        {
            if (settings == null) return null;
            var pinned = settings.Pinned;                 // ordered
            if (pinned == null) return null;
            if (slotIndex < 0 || slotIndex >= pinned.Count) return null;
            string tag = pinned[slotIndex];
            if (string.IsNullOrEmpty(tag) || registry == null) return null;
            foreach (var d in registry)
            {
                if (d != null && string.Equals(d.Tag, tag, StringComparison.Ordinal))
                    return d;
            }
            return null;   // catalogue miss
        }

        /// <summary>
        /// <c>getVisible</c> for a <c>slotLangNN</c> pinned-language slot button. A slot is visible
        /// iff there is a pinned language at that index and it resolves in the catalogue. Pure
        /// snapshot read (plus one locked registry lookup); deliberately does NOT call
        /// <c>_ribbon.Invalidate()</c> - re-entrant invalidate from inside a get* callback is a
        /// documented no-op / stall (see plan section 4 and Phase 0 findings). Never throws out of
        /// Office: on exception it logs and returns <c>false</c> (hidden).
        /// </summary>
        [System.CLSCompliant(false)]
        public bool GetSlotVisible(IRibbonControl control)
        {
            try
            {
                if (control == null) return false;
                int idx = SlotIndexFromControlId(control.Id);
                return ResolveSlot(idx) != null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: GetSlotVisible threw; hiding slot. " + ex);
                return false;
            }
        }

        /// <summary>
        /// <c>getLabel</c> for a <c>slotLangNN</c> slot button. Returns the resolved descriptor's
        /// label, falling back to its tag, then to <see cref="string.Empty"/> - never <c>null</c>.
        /// Pure snapshot read; does NOT self-invalidate. Never throws out of Office.
        /// </summary>
        [System.CLSCompliant(false)]
        public string GetSlotLabel(IRibbonControl control)
        {
            try
            {
                if (control == null) return string.Empty;
                int idx = SlotIndexFromControlId(control.Id);
                var desc = ResolveSlot(idx);
                if (desc == null) return string.Empty;
                return desc.Label ?? desc.Tag ?? string.Empty;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: GetSlotLabel threw; returning empty label. " + ex);
                return string.Empty;
            }
        }

        /// <summary>
        /// <c>getImage</c> for a <c>slotLangNN</c> slot button. Resolves the slot to a descriptor,
        /// loads the named bitmap from <c>Properties.Resources</c> and returns it as a COM
        /// <see cref="IStream"/> via <see cref="BuildImageStream"/> - the same proven mechanism
        /// the customUI <c>loadImage</c> callback (<see cref="GetImage(string)"/>) uses. On an
        /// unresolved slot (or any exception) it returns the <c>Other.png</c> image so a slot
        /// never shows a broken-icon glyph in the brief window between an invalidate and
        /// re-collection. Does NOT self-invalidate.
        /// </summary>
        [System.CLSCompliant(false)]
        public IStream GetSlotImage(IRibbonControl control)
        {
            try
            {
                string imageName = "Other.png";
                if (control != null)
                {
                    int idx = SlotIndexFromControlId(control.Id);
                    var desc = ResolveSlot(idx);
                    if (desc != null && !string.IsNullOrEmpty(desc.Image))
                        imageName = desc.Image;
                }

                return BuildImageStream(imageName) ?? BuildImageStream("Other.png");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: GetSlotImage threw; falling back to Other.png. " + ex);
                try
                {
                    return BuildImageStream("Other.png");
                }
                catch (Exception ex2)
                {
                    System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: GetSlotImage fallback also threw. " + ex2);
                    return null;
                }
            }
        }

        /// <summary>
        /// <c>getScreentip</c> for a <c>slotLangNN</c> slot button. Returns the resolved descriptor's
        /// screentip, or <see cref="string.Empty"/> - never <c>null</c> (reviewer M2: a
        /// <c>getScreentip</c> callback always runs and has no "omit" path; Office may treat a null
        /// return as a failed callback). <see cref="string.Empty"/> on both the unresolved path and
        /// the catch path. Pure snapshot read; does NOT self-invalidate.
        /// </summary>
        [System.CLSCompliant(false)]
        public string GetSlotScreentip(IRibbonControl control)
        {
            try
            {
                if (control == null) return string.Empty;
                int idx = SlotIndexFromControlId(control.Id);
                var desc = ResolveSlot(idx);
                if (desc == null) return string.Empty;
                return desc.Screentip ?? string.Empty;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: GetSlotScreentip threw; returning empty screentip. " + ex);
                return string.Empty;
            }
        }

        /// <summary>
        /// <c>getVisible</c> for the <c>menuMoreLanguages</c> dynamic menu. The menu hides
        /// when nothing is "enabled but not pinned" so the ribbon does not show an empty
        /// dropdown. Like <see cref="GetSlotVisible"/>, this callback is a pure
        /// snapshot read and must NOT call <c>_ribbon.Invalidate()</c> - Office is mid-
        /// collection of get* values and the call is re-entrant.
        /// </summary>
        [System.CLSCompliant(false)]
        public bool GetMoreMenuVisible(IRibbonControl control)
        {
            LanguageSettings snap;
            lock (_languagesLock) { snap = _languages; }
            if (snap == null) return false;

            // Cheap walk - Enabled is typically small (< 40 items) and we exit on the
            // first hit. Materialising a difference set per call would cost more.
            foreach (var tag in snap.Enabled)
            {
                if (!snap.IsPinned(tag)) return true;
            }
            return false;
        }

        /// <summary>
        /// <c>getContent</c> for the <c>menuMoreLanguages</c> dynamic menu. Returns a customUI
        /// menu fragment whose items are every enabled-but-not-pinned language declared in
        /// ribbon.xml, sorted by display label. Each item reuses <see cref="AddInButtonClicked"/>
        /// so the dynamic-menu click path is identical to a pinned-button click path.
        ///
        /// <para>
        /// The returned XML MUST declare the customUI namespace on the root &lt;menu&gt; element
        /// or Office silently drops the menu (plan section 4.2.1 / reviewer 2.9). Built via
        /// <see cref="XmlWriter"/> so labels containing &amp; / &lt; / &gt; / quotes are escaped
        /// rather than concatenated raw.
        /// </para>
        /// </summary>
        [System.CLSCompliant(false)]
        public string GetMoreLanguagesMenu(IRibbonControl control)
        {
            try
            {
                LanguageSettings snap;
                lock (_languagesLock) { snap = _languages; }
                return BuildMoreLanguagesMenuXml(snap, LanguageRegistry.All);
            }
            catch (Exception ex)
            {
                // Never let a get* callback throw out of Office - it tears down the ribbon for
                // the rest of the session. Log and return an empty (but well-formed) menu so
                // the dropdown opens silently.
                System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: GetMoreLanguagesMenu threw; returning empty menu. " + ex);
                return EmptyMoreLanguagesMenuXml();
            }
        }

        // The customUI namespace declared on the root <menu> element of the dynamic-menu
        // payload. Documented constant rather than an inline literal so tests can assert it.
        internal const string CustomUiNamespace = "http://schemas.microsoft.com/office/2006/01/customui";

        /// <summary>
        /// Build the dynamic-menu XML for <see cref="GetMoreLanguagesMenu"/>. Pure / static so it
        /// is exercised directly from unit tests without standing up an AddIn instance or an
        /// IRibbonControl proxy. See plan section 4.2.1 for the wire-shape contract.
        /// </summary>
        internal static string BuildMoreLanguagesMenuXml(LanguageSettings settings,
                                                         IReadOnlyList<LanguageDescriptor> registry)
        {
            // Build a tag->descriptor lookup once. Registry size is small (~30) so the
            // dictionary is negligible compared to the LINQ alternatives.
            var byTag = new Dictionary<string, LanguageDescriptor>(StringComparer.Ordinal);
            if (registry != null)
            {
                foreach (var d in registry)
                {
                    if (d == null || string.IsNullOrEmpty(d.Tag)) continue;
                    if (!byTag.ContainsKey(d.Tag)) byTag[d.Tag] = d;
                }
            }

            // Enabled-but-not-pinned, intersected with the registry (we cannot render a
            // descriptor we do not know about), sorted by display label using ordinal
            // case-insensitive ordering so "Go" / "go" do not swap on different locales.
            var items = new List<LanguageDescriptor>();
            if (settings != null)
            {
                foreach (var tag in settings.Enabled)
                {
                    if (settings.IsPinned(tag)) continue;
                    if (!byTag.TryGetValue(tag, out var d)) continue;
                    items.Add(d);
                }
            }
            items.Sort((a, b) => string.Compare(a.Label ?? a.Tag, b.Label ?? b.Tag, StringComparison.OrdinalIgnoreCase));

            var settingsXw = new XmlWriterSettings
            {
                OmitXmlDeclaration = true,
                ConformanceLevel = ConformanceLevel.Fragment,
            };

            using (var sw = new StringWriter(System.Globalization.CultureInfo.InvariantCulture))
            {
                using (var w = XmlWriter.Create(sw, settingsXw))
                {
                    w.WriteStartElement("menu", CustomUiNamespace);
                    foreach (var d in items)
                    {
                        w.WriteStartElement("button", CustomUiNamespace);
                        w.WriteAttributeString("id", "dyn_" + d.Tag);
                        w.WriteAttributeString("label", d.Label ?? d.Tag);
                        w.WriteAttributeString("tag", d.Tag);
                        w.WriteAttributeString("onAction", "AddInButtonClicked");
                        w.WriteAttributeString("image", string.IsNullOrEmpty(d.Image) ? "Other.png" : d.Image);
                        if (!string.IsNullOrEmpty(d.Screentip))
                            w.WriteAttributeString("screentip", d.Screentip);
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();
                }
                return sw.ToString();
            }
        }

        private static string EmptyMoreLanguagesMenuXml()
        {
            return "<menu xmlns=\"" + CustomUiNamespace + "\"/>";
        }

        /// <summary>
        /// Common pre-amble for every <c>onAction</c> ribbon callback. If a worker has flipped
        /// <see cref="_invalidatePending"/> since the last click, this is the only place we call
        /// <see cref="IRibbonUI.Invalidate"/>. The flag is read-then-cleared without an interlocked
        /// because Office serialises onAction delivery per ribbon control and the worst case is a
        /// redundant Invalidate.
        /// </summary>
        private void ObservePendingInvalidate()
        {
            if (!_invalidatePending) return;
            _invalidatePending = false;
            try
            {
                _ribbon?.Invalidate();
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                // Host is shutting down or RPC failed - swallow. Settings are already persisted.
                System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: _ribbon.Invalidate() threw COMException; ignoring. " + ex.Message);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: _ribbon.Invalidate() threw; ignoring. " + ex.Message);
            }
        }

        public bool cbQuickStyle_GetPressed(IRibbonControl control)
        {
            this.QuickStyle = NoteHighlightForm.Properties.Settings.Default.QuickStyle;
            return this.QuickStyle;
        }

        public void cbQuickStyle_OnAction(IRibbonControl control, bool isPressed)
        {
            ObservePendingInvalidate();
            this.QuickStyle = isPressed;
            NoteHighlightForm.Properties.Settings.Default.QuickStyle = this.QuickStyle;
            SettingsHelper.SafeSave();
        }


        public bool cbDarkMode_GetPressed(IRibbonControl control)
        {
            this.DarkMode = NoteHighlightForm.Properties.Settings.Default.DarkMode;
            return this.DarkMode;
        }

        public void cbDarkMode_OnAction(IRibbonControl control, bool isPressed)
        {
            ObservePendingInvalidate();
            this.DarkMode = isPressed;
            NoteHighlightForm.Properties.Settings.Default.DarkMode = this.DarkMode;
            SettingsHelper.SafeSave();
        }

        /// <summary>
        /// Lazily construct the main-worker filter+thread on first use. Called only from the main
        /// STA (ribbon callback), which Office serialises - a plain null-check is therefore enough.
        /// </summary>
        private StaWorker EnsureMainWorker()
        {
            if (_mainWorker == null)
            {
                _mainFilter = new OneNoteMessageFilter();
                _mainWorker = new StaWorker("Main", _mainFilter);
                _mainWorker.Start();
            }
            return _mainWorker;
        }

        /// <summary>
        /// Lazily construct the settings-worker filter+thread on first use. See
        /// <see cref="EnsureMainWorker"/> for the lock-free rationale.
        /// </summary>
        private StaWorker EnsureSettingsWorker()
        {
            if (_settingsWorker == null)
            {
                _settingsFilter = new OneNoteMessageFilter();
                _settingsWorker = new StaWorker("Settings", _settingsFilter);
                _settingsWorker.Start();
            }
            return _settingsWorker;
        }

        public void AddInButtonClicked(IRibbonControl control)
        {
            ObservePendingInvalidate();
            try
            {
                // Capture control.Tag (an immutable, apartment-safe managed string) on the
                // ribbon-callback thread BEFORE crossing apartments. DO NOT capture `control`
                // itself - IRibbonControl is an RCW owned by Office's ribbon dispatcher and
                // must not cross apartments.
                string clickTag = control.Tag;
                EnsureMainWorker().Post(() => ShowForm(clickTag));
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception from AddInButtonClicked: " + e.ToString());
            }
        }

        /// <summary>
        /// <c>onAction</c> for a <c>slotLangNN</c> pinned-language slot button. Unlike
        /// <see cref="AddInButtonClicked"/> - which is still used by the dynamic-menu items and
        /// maps by <c>control.Tag</c> - slot buttons carry no <c>tag=</c> attribute, so this maps
        /// by <c>control.Id</c> via <see cref="SlotIndexFromControlId"/> and resolves the tag at
        /// click time. The resolved <c>desc.Tag</c> (an immutable, apartment-safe managed string)
        /// is captured BEFORE the cross-apartment Post, exactly as <see cref="AddInButtonClicked"/>
        /// captures <c>control.Tag</c>.
        /// </summary>
        [System.CLSCompliant(false)]
        public void SlotButtonClicked(IRibbonControl control)
        {
            ObservePendingInvalidate();
            try
            {
                int idx = SlotIndexFromControlId(control?.Id);   // map by Id, not Tag
                var desc = ResolveSlot(idx);
                if (desc == null)
                {
                    System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: slot click on unresolved slot id='"
                        + (control?.Id ?? "<null>") + "'.");
                    return;
                }
                string clickTag = desc.Tag;          // immutable managed string
                EnsureMainWorker().Post(() => ShowForm(clickTag));
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception from SlotButtonClicked: " + e.ToString());
            }
        }

        private void ShowForm(string tag)
        {
            string outFileName = Guid.NewGuid().ToString();
            string htmlOutputPath = Path.Combine(Path.GetTempPath(), outFileName + ".html");

            try
            {
                var pageNode = GetPageNode();
                XElement pageRoot = null;
                string selectedText = "";
                XElement outline = null;
                bool selectedTextFormated = false;

                PageSelection selection = PageSelection.From(null, ns);
                if (pageNode != null)
                {
                    string pageXml = GetPageXml(pageNode.Attribute("ID").Value);

                    // H9: parse the page XML once and share the parsed root with every helper that
                    // needs it. OneNote can return zero-length XML or HTML error fragments mid-sync,
                    // so a parse failure aborts the operation silently rather than bubbling a stack
                    // trace through a MessageBox on this worker STA.
                    try
                    {
                        pageRoot = XDocument.Parse(pageXml).Root;
                    }
                    catch (System.Xml.XmlException xex)
                    {
                        System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: could not parse page XML in ShowForm; aborting. " + xex.Message);
                        return;
                    }

                    // Resolve the selection state once and thread it through every helper that
                    // would otherwise re-traverse the page (resolves review item 2.1).
                    selection = PageSelection.From(pageRoot, ns);
                    selectedText = GetSelectedText(selection, out selectedTextFormated);

                    if (selectedText.Trim() != "")
                    {
                        outline = selection.Outline;
                    }
                }

                MainForm form;
                try
                {
                    form = new MainForm(tag, outFileName, selectedText, this.QuickStyle, this.DarkMode);
                }
                catch (Exception ex)
                {
                    // Surface constructor failures to the user. The worker is pumping its own
                    // message loop while this action runs, so MessageBox is safe here.
                    MessageBox.Show("Could not open NoteHighlight: " + ex.Message);
                    return;
                }

                _currentMainForm = form;
                try
                {
                    // Application.Run is required, NOT form.ShowDialog(). ShowDialog needs an
                    // owner-window context this worker STA does not have, and its modal semantics
                    // interact poorly with the absence of a parent form on the worker. See plan
                    // section 3 ("Click coalescing - why Application.Run(form) serialises forms").
                    System.Windows.Forms.Application.Run(form);
                }
                finally
                {
                    _currentMainForm = null;
                }

                if (File.Exists(htmlOutputPath))
                {
                    InsertHighLightCodeToCurrentSide(htmlOutputPath, selection, form.Parameters, outline, selectedTextFormated);
                }
            }
            catch (Exception e)
            {
                // H9: do not surface the raw stack trace via MessageBox from the worker STA - that
                // blocks the worker's message pump and leaves the OneNote ribbon spinning. Log it.
                System.Diagnostics.Trace.TraceError("NoteHighlight2016: unhandled exception in ShowForm. " + e.ToString());
            }
            finally
            {
                // Best-effort temp-file cleanup. Always attempt to delete the highlighter output
                // even if InsertHighLightCodeToCurrentSide threw or the user cancelled the form.
                try
                {
                    if (File.Exists(htmlOutputPath))
                    {
                        File.Delete(htmlOutputPath);
                    }
                }
                catch
                {
                    // Best-effort: do not surface temp-file delete failures to the user.
                }
            }
        }

        public void SettingsButtonClicked(IRibbonControl control)
        {
            ObservePendingInvalidate();
            try
            {
                EnsureSettingsWorker().Post(() => ShowSettingsForm());
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception from SettingsButtonClicked: " + e.ToString());
            }
        }

        private void ShowSettingsForm()
        {
            try
            {
                SettingsForm form;
                try
                {
                    form = new SettingsForm();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not open NoteHighlight settings: " + ex.Message);
                    return;
                }

                _currentSettingsForm = form;
                try
                {
                    System.Windows.Forms.Application.Run(form);
                }
                finally
                {
                    _currentSettingsForm = null;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception from ShowSettingsForm: " + e.ToString());
            }
        }

        /// <summary>
        /// onAction for the "Languages..." ribbon button in the Advanced group. Dispatches to the
        /// existing settings STA worker (mirrors <see cref="SettingsButtonClicked"/>). The actual
        /// form (LanguageSettingsForm) is constructed and Run on the worker thread; this method
        /// must return promptly so OneNote's ribbon dispatcher does not stall.
        /// </summary>
        [System.CLSCompliant(false)]
        public void LanguagesButtonClicked(IRibbonControl control)
        {
            ObservePendingInvalidate();
            try
            {
                // Snapshot the current LanguageSettings under the lock and hand the worker a deep
                // copy so the user can edit independently of the live ribbon state. The save path
                // calls ReplaceLanguageSettings(...) with the post-edit instance.
                LanguageSettings snap;
                lock (_languagesLock) { snap = _languages; }

                // Defensive: if OnConnection's seed somehow failed we still want the dialog to
                // open, so the user can hit Reset to defaults and recover.
                var editable = LanguageSettingsForm.CloneForEditing(snap, LanguageRegistry.DefaultPinned);
                var registry = LanguageRegistry.All;
                var jsonPath = SettingsHelper.LanguagesJsonPath;

                EnsureSettingsWorker().Post(() => ShowLanguageSettingsForm(editable, registry, jsonPath));
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception from LanguagesButtonClicked: " + e.ToString());
            }
        }

        private void ShowLanguageSettingsForm(LanguageSettings editable,
                                              IReadOnlyList<LanguageDescriptor> registry,
                                              string jsonPath)
        {
            try
            {
                LanguageSettingsForm form;
                try
                {
                    form = new LanguageSettingsForm(this, editable, registry, jsonPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not open NoteHighlight Languages dialog: " + ex.Message);
                    return;
                }

                _currentLanguagesForm = form;
                try
                {
                    System.Windows.Forms.Application.Run(form);
                }
                finally
                {
                    _currentLanguagesForm = null;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception from ShowLanguageSettingsForm: " + e.ToString());
            }
        }

        /// <summary>
        /// Called by <see cref="LanguageSettingsForm"/> on the settings STA after a successful
        /// save. Atomically swaps the live <c>_languages</c> reference under <see cref="_languagesLock"/>
        /// and sets <see cref="_invalidatePending"/> so the next ribbon onAction triggers an
        /// IRibbonUI.Invalidate(). DOES NOT touch <see cref="_ribbon"/> directly - the worker has
        /// no apartment-safe handle on Office's ribbon-callback threads.
        ///
        /// <para>
        /// Thread-safety invariant (reviewer m1): this method must SWAP the whole <c>_languages</c>
        /// reference, never mutate the existing instance in place. <see cref="ResolveSlot(int)"/>
        /// and the slot get* callbacks read <c>snap.Pinned</c> lock-free after taking a snapshot
        /// under <see cref="_languagesLock"/>; that is only safe because the live instance is
        /// immutable once published. <see cref="LanguageSettings"/> does expose mutators
        /// (Pin/Unpin/Reorder) - a future change here that called one of those on the live
        /// <c>_languages</c> instead of swapping a freshly built one would quietly break that
        /// invariant.
        /// </para>
        /// </summary>
        internal void ReplaceLanguageSettings(LanguageSettings settings)
        {
            if (settings == null) return;
            lock (_languagesLock)
            {
                _languages = settings;
            }
            _invalidatePending = true;
        }

        /// <summary>
        /// Reflects over <c>Properties.Resources</c> by name and returns the matching
        /// <see cref="Bitmap"/>, or <c>null</c> if the resource is missing or is not a Bitmap.
        /// Shared lookup core for both image callbacks: <see cref="GetImage(string)"/>
        /// (the customUI <c>loadImage</c> path) and <see cref="GetSlotImage"/> (the
        /// per-control <c>getImage</c> path). Both wrap the result as a COM <c>IStream</c>
        /// via <see cref="BuildImageStream"/>.
        ///
        /// <para>
        /// Memory rule project_ribbon_icons_load_via_resx_reflection still applies: every image
        /// referenced by the ribbon needs a ResX &lt;data&gt; entry plus a Designer.cs accessor,
        /// not merely a csproj &lt;Content&gt; row - the lookup here is reflection over the
        /// generated <c>Properties.Resources</c> accessors.
        /// </para>
        /// </summary>
        private static Bitmap LoadImageBitmapByName(string imageName)
        {
            if (string.IsNullOrEmpty(imageName)) return null;

            BindingFlags flags = BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

            int dot = imageName.IndexOf('.');
            string propertyName = dot >= 0 ? imageName.Substring(0, dot) : imageName;

            PropertyInfo prop = typeof(Properties.Resources).GetProperty(propertyName, flags);
            if (prop == null)
            {
                System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: ribbon image resource '" + propertyName + "' not found.");
                return null;
            }

            Bitmap b = prop.GetValue(null, null) as Bitmap;
            if (b == null)
            {
                System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: ribbon image resource '" + propertyName + "' is not a Bitmap.");
                return null;
            }
            return b;
        }

        /// <summary>
        /// Specified in Ribbon.xml as the root <c>loadImage</c> callback, this method returns the
        /// image to display on a ribbon button declared with a static <c>image=</c> attribute.
        /// Returns a COM <c>IStream</c> via <see cref="BuildImageStream"/> - the same mechanism
        /// <see cref="GetSlotImage"/> (the per-control <c>getImage</c> callback) uses.
        /// </summary>
        /// <param name="imageName"></param>
        /// <returns></returns>
        public IStream GetImage(string imageName)
		{
            return BuildImageStream(imageName);
		}

        /// <summary>
        /// Loads a named resource bitmap and serialises it into a COM <see cref="IStream"/>
        /// (PNG bytes, position rewound to 0). Shared by <see cref="GetImage(string)"/> (the
        /// customUI <c>loadImage</c> path) and <see cref="GetSlotImage"/> (the per-control
        /// <c>getImage</c> path).
        ///
        /// <para>
        /// Both ribbon image callbacks in this add-in return <c>IStream</c>. The Office ribbon
        /// in this shared-add-in host consumes an <c>IStream</c> from a <c>getImage</c> callback
        /// just as it does from <c>loadImage</c>; the earlier <c>stdole.IPictureDisp</c> route
        /// (via an <c>AxHost</c> shim) produced a picture Office would not bind, so every slot
        /// button rendered blank. Returning the same <c>CCOMStreamWrapper</c> the known-good
        /// <c>loadImage</c> path uses keeps both callbacks on one proven mechanism.
        /// </para>
        ///
        /// Returns <c>null</c> if the resource is missing (e.g. a button references an image
        /// that was not embedded) so Office falls back to a default icon rather than crashing
        /// the whole ribbon with "An error occurred while creating the ribbon".
        /// </summary>
        private static IStream BuildImageStream(string imageName)
        {
            MemoryStream imageStream = new MemoryStream();

            // Dispose the source Bitmap deterministically after Save - the PNG bytes have
            // already been serialised into the MemoryStream, so disposing the bitmap does
            // not affect the stream content.
            using (Bitmap b = LoadImageBitmapByName(imageName))
            {
                if (b == null)
                {
                    imageStream.Dispose();
                    return null;
                }
                b.Save(imageStream, ImageFormat.Png);
            }

            // Bitmap.Save leaves Position at end-of-stream; rewind so callers that read from
            // the current position (rather than Seek to 0) receive the PNG.
            imageStream.Position = 0;

            return new CCOMStreamWrapper(imageStream);
        }

        /// <summary>
        /// Insert HighLight Code To Mouse Position.
        /// </summary>
        private void InsertHighLightCodeToCurrentSide(string fileName, PageSelection selection, HighLightParameter parameters, XElement outline, bool selectedTextFormated)
        {
            try
            {
                string htmlContent = File.ReadAllText(fileName, new UTF8Encoding(false));

                string byteOrderMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());
                htmlContent = htmlContent.Replace(byteOrderMarkUtf8, "");

                var pageNode = GetPageNode();

                if (pageNode != null)
                {
                    var existingPageId = pageNode.Attribute("ID").Value;
                    string[] position = null;
                    if (outline == null)
                    {
                        position = GetMousePointPosition(selection);
                    }

                    var page = InsertHighLightCode(htmlContent, position, parameters, outline, (new GenerateHighLight(GetAddinDirectory())).Config, selectedTextFormated, IsSelectedTextInline(selection));
                    page.Root.SetAttributeValue("ID", existingPageId);

                    //Bug fix - remove overflow value for Indents
                    foreach (var el in page.Descendants(ns + "Indent").Where(n => n.Attribute("indent") != null && double.TryParse(n.Attribute("indent").Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) && d > 1e6))
                    {
                        el.Attribute("indent").Value = "0";
                    }

                    OneNoteApplication.UpdatePageContent(page.ToString(), DateTime.MinValue);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception from InsertHighLightCodeToCurrentSide: "+e.ToString());
            }
        }

        XElement GetPageNode()
        {
            string notebookXml;
            try
            {
                OneNoteApplication.GetHierarchy(null, HierarchyScope.hsPages, out notebookXml, XMLSchema.xs2013);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception from onApp.GetHierarchy:" + ex.Message);
                return null;
            }

            var doc = XDocument.Parse(notebookXml);
            ns = doc.Root.Name.Namespace;

            var pageNode = doc.Descendants(ns + "Page")
                              .Where(n => n.Attribute("isCurrentlyViewed") != null && n.Attribute("isCurrentlyViewed").Value == "true")
                              .FirstOrDefault();
            return pageNode;
        }

        /// <summary>
        /// Get Mouse Point. Consumes the pre-resolved <see cref="PageSelection"/>
        /// rather than re-traversing the page XML (resolves review item 2.1).
        /// </summary>
        private string[] GetMousePointPosition(PageSelection selection)
        {
            if (selection == null) return null;
            var node = selection.PartialOutline;
            if (node != null)
            {
                var attrPos = node.Descendants(ns + "Position").FirstOrDefault();
                if (attrPos != null)
                {
                    var x = attrPos.Attribute("x").Value;
                    var y = attrPos.Attribute("y").Value;
                    return new string[] { x, y };
                }
            }
            return null;
        }

        private string GetPageXml(string pageID)
        {
            string pageXml;
            OneNoteApplication.GetPageContent(pageID, out pageXml, PageInfo.piSelection);

            return pageXml;
        }

        /// <summary>
        /// Test-seam overload kept for backwards compatibility with the
        /// existing UnitTesting fixtures that pass a raw <c>pageRoot</c>.
        /// New callers should build a <see cref="PageSelection"/> once via
        /// <see cref="PageSelection.From"/> and pass that through instead.
        /// </summary>
        public string GetSelectedText(XElement pageRoot, out bool selectedTextFormated)
        {
            return GetSelectedText(PageSelection.From(pageRoot, ns), out selectedTextFormated);
        }

        internal string GetSelectedText(PageSelection selection, out bool selectedTextFormated)
        {
            selectedTextFormated = false;
            StringBuilder sb = new StringBuilder();
            if (selection == null || selection.Outline == null) return sb.ToString();

            var node = selection.Outline;
            var table = selection.Table;

            System.Collections.Generic.IEnumerable<XElement> attrPos;
            if (table == null)
            {
                attrPos = node.Descendants(ns + "T").Where(n => n.Attribute("selected") != null && n.Attribute("selected").Value == "all");
            }
            else
            {
                attrPos = table.Descendants(ns + "Cell")
                               .SelectMany(c => c.Descendants(ns + "T"))
                               .Where(n => n.Attribute("selected")?.Value == "all");
                selectedTextFormated = true;
            }
            int tabCount = 0;
            int initTabCount = -1;
            foreach (var line in attrPos)
            {
                var htmlDocument = new HtmlAgilityPack.HtmlDocument();
                htmlDocument.LoadHtml(line.Value);

                if (initTabCount == -1)
                {
                    initTabCount = line.Ancestors().Elements(ns + "T").Count();
                }
                tabCount = line.Ancestors().Elements(ns + "T").Count() - initTabCount;


                sb.AppendLine(new String('\t', tabCount) + HttpUtility.HtmlDecode(htmlDocument.DocumentNode.InnerText));
            }
            return sb.ToString().TrimEnd('\r','\n');
        }

        /// <summary>
        /// Test-seam overload kept for backwards compatibility with the
        /// existing UnitTesting fixtures that pass a raw <c>pageRoot</c>.
        /// </summary>
        public bool IsSelectedTextInline(XElement pageRoot)
        {
            return IsSelectedTextInline(PageSelection.From(pageRoot, ns));
        }

        internal bool IsSelectedTextInline(PageSelection selection)
        {
            if (selection == null || selection.Outline == null) return false;

            var node = selection.Outline;
            var table = selection.Table;
            var scope = table ?? node;

            foreach (var oeNode in scope.Descendants(ns + "OE"))
            {
                var allSel = oeNode.Descendants(ns + "T").Any(n => n.Attribute("selected") != null && n.Attribute("selected").Value == "all");
                var unsel  = oeNode.Descendants(ns + "T").Any(n => n.Attribute("selected") == null || n.Attribute("selected").Value == "none");
                if (allSel && unsel) return true;
            }
            return false;
        }

        /// <summary>
        /// Generate XML Insert To OneNote.
        /// </summary>
        public XDocument InsertHighLightCode(string htmlContent, string[] position, HighLightParameter parameters, XElement outline, HighLightSection config, bool selectedTextFormated, bool isInline)
        {
            XElement children = PrepareFormatedContent(htmlContent, parameters, config, isInline);

            bool update = false;
            if (outline == null)
            {
                outline = CreateOutline(position, children);
            }
            else // Update exiting outline
            {
                update = true;

                //Change outline width
                outline.Element(ns + "Size").Attribute("width").Value = "1600";

                if (selectedTextFormated)
                {
                    outline.Descendants(ns + "Table").Where(n => n.Attribute("selected") != null &&
                                        (n.Attribute("selected").Value == "all" || n.Attribute("selected").Value == "partial")).FirstOrDefault().ReplaceWith(children.Descendants(ns + "Table").FirstOrDefault());
                }
                else
                {
                    if (isInline)
                    {
                        int j = 0;
                        for(int i = 0; i < outline.Descendants(ns + "OE").Count(); i++)
                        {
                            XElement oeNode = outline.Descendants(ns + "OE").ElementAt(i);

                            if (oeNode.Descendants(ns + "T").Where(n => n.Attribute("selected") != null && n.Attribute("selected").Value == "all").Count() > 0)
                            {
                                oeNode.Descendants(ns + "T").Where(n => n.Attribute("selected") != null && n.Attribute("selected").Value == "all").FirstOrDefault().ReplaceWith(children.Descendants(ns + "Table").Descendants(ns + "OEChildren").Descendants(ns + "OE").ElementAt(j).Descendants(ns + "T"));
                                j++;
                            }

                        }
                        outline.Descendants(ns + "OE").Where(t => t.Elements(ns + "T").Any(n => n.Attribute("selected") != null && n.Attribute("selected").Value == "all")).Remove();
                    }
                    else
                    {
                        outline.Descendants(ns + "T").Where(n => n.Attribute("selected") != null && n.Attribute("selected").Value == "all").FirstOrDefault().ReplaceWith(children.Descendants(ns + "Table").FirstOrDefault());
                        outline.Descendants(ns + "OE").Where(t => t.Elements(ns + "T").Any(n => n.Attribute("selected") != null && n.Attribute("selected").Value == "all")).Remove();
                        outline.Descendants(ns + "OEChildren").Where(n => n.HasElements == false && n.Attribute("selected") != null && (n.Attribute("selected").Value == "partial")).Remove();
                    }
                }
            }

            if (update)
            {
                return outline.Parent.Document;
            }
            else
            {
                XElement page = new XElement(ns + "Page");
                page.Add(outline);

                XDocument doc = new XDocument();
                doc.Add(page);
                return doc;
            }


        }

        private XElement CreateOutline(string[] position, XElement children)
        {
            XElement outline = new XElement(ns + "Outline");
            if (position != null && position.Length == 2)
            {
                XElement pos = new XElement(ns + "Position");
                pos.Add(new XAttribute("x", position[0]));
                pos.Add(new XAttribute("y", position[1]));
                outline.Add(pos);

                XElement size = new XElement(ns + "Size");
                size.Add(new XAttribute("width", "1600"));
                size.Add(new XAttribute("height", "200"));
                outline.Add(size);
            }
            outline.Add(children);
            return outline;
        }

        private XElement PrepareFormatedContent(string htmlContent, HighLightParameter parameters, HighLightSection config, bool isInline)
        {
            XElement children = new XElement(ns + "OEChildren");

            XElement table = new XElement(ns + "Table");
            table.Add(new XAttribute("bordersVisible", NoteHighlightForm.Properties.Settings.Default.ShowTableBorder));

            XElement columns = new XElement(ns + "Columns");
            XElement column1 = new XElement(ns + "Column");
            column1.Add(new XAttribute("index", "0"));
            column1.Add(new XAttribute("width", "40"));
            if (parameters.ShowLineNumber && !isInline)
            {
                columns.Add(column1);
            }
            XElement column2 = new XElement(ns + "Column");
            if (parameters.ShowLineNumber && !isInline)
            {
                column2.Add(new XAttribute("index", "1"));
            }
            else
            {
                column2.Add(new XAttribute("index", "0"));
            }

            column2.Add(new XAttribute("width", "1400"));
            columns.Add(column2);

            table.Add(columns);

            Color color = parameters.HighlightColor;
            string colorString = color.A == 0 ? "none" : string.Format("#{0:X2}{1:X2}{2:X2}", color.R, color.G, color.B);

            XElement row = new XElement(ns + "Row");
            XElement cell1 = new XElement(ns + "Cell");
            cell1.Add(new XAttribute("shadingColor", colorString));
            XElement cell2 = new XElement(ns + "Cell");
            cell2.Add(new XAttribute("shadingColor", colorString));


            string defaultStyle = "";

            var arrayLine = htmlContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            foreach (var it in arrayLine)
            {
                string item = it;

                if (item.StartsWith("<pre"))
                {
                    defaultStyle = item.Substring(0, item.IndexOf(">") + 1);
                    //Sets language to Latin to disable spell check
                    defaultStyle = defaultStyle.Insert(defaultStyle.Length - 1, " lang=la");

                    if (this.DarkMode)
                    {
                        //Remove background-color element so that text would render with correct contrast in dark mode
                        defaultStyle = PreviewHtmlWrapper.StripPreBackgroundColor(defaultStyle);
                    }

                    item = item.Substring(item.IndexOf(">") + 1);
                }

                if (item == "</pre>")
                {
                    continue;
                }

                var itemNr = "";
                var itemLine = "";
                if (parameters.ShowLineNumber)
                {
                    if (item.Contains("</span>"))
                    {
                        int ind = item.IndexOf("</span>");
                        itemNr = item.Substring(0, ind + ("</span>").Length);
                        itemLine = item.Substring(ind);
                    }
                    else
                    {
                        itemNr = "";
                        itemLine = item;
                    }

                    string nr = "";
                    if (string.IsNullOrEmpty(config.LineNrReplaceCh))
                    {
                        nr = defaultStyle + itemNr.Replace("&apos;", "'") + "</pre>";
                    }
                    else
                    {
                        nr = defaultStyle + config.LineNrReplaceCh.PadLeft(5) + "</pre>";
                    }

                    XElement oeElement = new XElement(ns + "OE",
                                    new XElement(ns + "T",
                                        new XCData(nr)));
                    if (ContainsAsianCharacter(itemLine))
                    {
                        oeElement.Add(new XAttribute("spaceBefore", config.AsianBeforeSpace));
                        oeElement.Add(new XAttribute("spaceAfter", config.AsianAfterSpace));
                    }

                    cell1.Add(new XElement(ns + "OEChildren",
                               oeElement ));
                }
                else
                {
                    itemLine = item;
                }
                string s = defaultStyle + itemLine.Replace("&apos;", "'") + "</pre>";

                cell2.Add(new XElement(ns + "OEChildren",
                            new XElement(ns + "OE",
                                new XElement(ns + "T",
                                    new XCData(s)))));

            }

            if (parameters.ShowLineNumber && !isInline)
            {
                row.Add(cell1);
            }
            row.Add(cell2);

            table.Add(row);

            children.Add(new XElement(ns + "OE",
                                table));
            return children;
        }

        private bool ContainsAsianCharacter(string itemLine)
        {
            return itemLine.Any(c => (uint)c >= 0x4E00 && (uint)c <= 0x2FA1F);
        }
    }
}
