/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 */

using System;
using System.Diagnostics.CodeAnalysis;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Extensibility;
using Microsoft.Office.Core;
using NoteHighlightAddin.Utilities;
using Application = Microsoft.Office.Interop.OneNote.Application;  // Conflicts with System.Windows.Forms
using System.Reflection;
using System.Drawing;
using Microsoft.Office.Interop.OneNote;
using NoteHighLightForm;
using System.Text;
using System.Linq;
using Helper;
using System.Threading;
using System.Web;
using GenerateHighlightContent;
using System.Configuration;
using System.Globalization;

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

        // Filters are held by the AddIn so SignalShutdown() can be invoked from OnBeginShutdown /
        // OnDisconnection on the main STA. The filter's _shuttingDown flag is volatile, so this
        // cross-thread write is safe without additional synchronisation.
        private OneNoteMessageFilter _mainFilter;
        private OneNoteMessageFilter _settingsFilter;

        private bool QuickStyle { get; set; }

        private bool DarkMode { get; set; }

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
            try
            {

                // Use the defining assembly (typeof(AddIn).Assembly) rather than
                // GetCallingAssembly(): the latter resolves to whichever assembly
                // happens to call GetCustomUI - under cross-AppDomain COM activation
                // that is mscorlib (shadow-copied path) or ONENOTE.EXE itself, neither
                // of which sit next to ribbon.xml. The current call only works because
                // the JIT inlines this helper into the COM entry point.
                var assemblyLocation = typeof(AddIn).Assembly.Location;
                if (string.IsNullOrEmpty(assemblyLocation))
                    assemblyLocation = new Uri(typeof(AddIn).Assembly.CodeBase).LocalPath;
                var workingDirectory = Path.Combine(ProcessHelper.GetDirectoryFromPath(assemblyLocation), "ribbon.xml");

                string file = File.ReadAllText(workingDirectory);

                return file;

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
			SetOneNoteApplication((Application)Application);
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

        public bool cbQuickStyle_GetPressed(IRibbonControl control)
        {
            this.QuickStyle = NoteHighlightForm.Properties.Settings.Default.QuickStyle;
            return this.QuickStyle;
        }

        public void cbQuickStyle_OnAction(IRibbonControl control, bool isPressed)
        {
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
            try
            {
                // Capture control.Tag (an immutable, apartment-safe managed string) on the main STA
                // BEFORE crossing apartments. DO NOT capture `control` itself - IRibbonControl is an
                // RCW belonging to OneNote's main STA and must not cross apartments.
                string clickTag = control.Tag;
                EnsureMainWorker().Post(() => ShowForm(clickTag));
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception from AddInButtonClicked: " + e.ToString());
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

                    selectedText = GetSelectedText(pageRoot, out selectedTextFormated);

                    if (selectedText.Trim() != "")
                    {
                        outline = GetOutline(pageRoot);
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
                    InsertHighLightCodeToCurrentSide(htmlOutputPath, pageRoot, form.Parameters, outline, selectedTextFormated);
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
        /// Specified in Ribbon.xml, this method returns the image to display on the ribbon button
        /// </summary>
        /// <param name="imageName"></param>
        /// <returns></returns>
        public IStream GetImage(string imageName)
		{
            //switch (imageName)
            //{
            //    case "CSharp.png":
            //        Properties.Resources.CSharp.Save(imageStream, ImageFormat.Png);
            //        break;
            //    default:
            //        Properties.Resources.Logo.Save(imageStream, ImageFormat.Png);
            //        break;
            //}

            BindingFlags flags = BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

            // H7: null-guard the reflected resource lookup. If the resource is missing
            // (e.g. someone added a button referencing an image that was not embedded),
            // return null so Office falls back to a default icon rather than crashing the
            // whole ribbon with "An error occurred while creating the ribbon".
            string propertyName = imageName.Substring(0, imageName.IndexOf('.'));
            PropertyInfo prop = typeof(Properties.Resources).GetProperty(propertyName, flags);
            if (prop == null)
            {
                System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: ribbon image resource '" + propertyName + "' not found.");
                return null;
            }

            // H7: dispose the source Bitmap deterministically after Save - the PNG bytes
            // have already been serialised into the MemoryStream, so disposing the bitmap
            // does not affect the stream content.
            MemoryStream imageStream = new MemoryStream();
            using (Bitmap b = prop.GetValue(null, null) as Bitmap)
            {
                if (b == null)
                {
                    System.Diagnostics.Trace.TraceWarning("NoteHighlight2016: ribbon image resource '" + propertyName + "' is not a Bitmap.");
                    imageStream.Dispose();
                    return null;
                }
                b.Save(imageStream, ImageFormat.Png);
            }

            // H7: Bitmap.Save leaves Position at end-of-stream; rewind so callers that
            // read from the current position (rather than Seek to 0) receive the PNG.
            imageStream.Position = 0;

            return new CCOMStreamWrapper(imageStream);
		}

        /// <summary>
        /// 插入 HighLight Code 至滑鼠游標的位置
        /// Insert HighLight Code To Mouse Position  
        /// </summary>
        private void InsertHighLightCodeToCurrentSide(string fileName, XElement pageRoot, HighLightParameter parameters, XElement outline, bool selectedTextFormated)
        {
            try
            {
                // Trace.TraceInformation(System.Reflection.MethodBase.GetCurrentMethod().Name);
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
                        position = GetMousePointPosition(pageRoot);
                    }

                    var page = InsertHighLightCode(htmlContent, position, parameters, outline, (new GenerateHighLight(GetAddinDirectory())).Config, selectedTextFormated, IsSelectedTextInline(pageRoot));
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
                return null; ;
            }

            var doc = XDocument.Parse(notebookXml);
            ns = doc.Root.Name.Namespace;

            var pageNode = doc.Descendants(ns + "Page")
                              .Where(n => n.Attribute("isCurrentlyViewed") != null && n.Attribute("isCurrentlyViewed").Value == "true")
                              .FirstOrDefault();
            return pageNode;
        }

        /// <summary>
        /// 取得滑鼠所在的點
        /// Get Mouse Point
        /// </summary>
        private string[] GetMousePointPosition(XElement pageRoot)
        {
            if (pageRoot == null) return null;
            var node = pageRoot.Descendants(ns + "Outline")
                               .Where(n => n.Attribute("selected") != null && n.Attribute("selected").Value == "partial")
                               .FirstOrDefault();
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

        private XElement GetOutline(XElement pageRoot)
        {
            if (pageRoot == null) return null;
            var node = pageRoot.Descendants(ns + "Outline")
                               .Where(n => n.Attribute("selected") != null && (n.Attribute("selected").Value == "all" || n.Attribute("selected").Value == "partial"))
                               .FirstOrDefault();
            //if (node != null)
            //{
            //    var attrPos = node.Descendants(ns + "Position").FirstOrDefault();
            //    if (attrPos != null)
            //    {
            //        var x = attrPos.Attribute("x").Value;
            //        var y = attrPos.Attribute("y").Value;
            //        return new string[] { x, y };
            //    }
            //}
            //return null;

            return node;
        }

        private string GetPageXml(string pageID)
        {
            string pageXml;
            OneNoteApplication.GetPageContent(pageID, out pageXml, PageInfo.piSelection);

            return pageXml;
        }

        public string GetSelectedText(XElement pageRoot, out bool selectedTextFormated)
        {
            selectedTextFormated = false;
            StringBuilder sb = new StringBuilder();
            if (pageRoot == null) return sb.ToString();
            var node = pageRoot.Descendants(ns + "Outline")
                               .Where(n => n.Attribute("selected") != null && (n.Attribute("selected").Value == "all" || n.Attribute("selected").Value == "partial"))
                               .FirstOrDefault();

            if (node != null)
            {
                var table = node.Descendants(ns + "Table").Where(n => n.Attribute("selected") != null && (n.Attribute("selected").Value == "all" || n.Attribute("selected").Value == "partial")).FirstOrDefault();

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
            }
            return sb.ToString().TrimEnd('\r','\n');
        }

        public bool IsSelectedTextInline(XElement pageRoot)
        {
            if (pageRoot == null) return false;
            var node = pageRoot.Descendants(ns + "Outline")
                               .Where(n => n.Attribute("selected") != null && (n.Attribute("selected").Value == "all" || n.Attribute("selected").Value == "partial"))
                               .FirstOrDefault();

            if (node != null)
            {
                var table = node.Descendants(ns + "Table").Where(n => n.Attribute("selected") != null && (n.Attribute("selected").Value == "all" || n.Attribute("selected").Value == "partial")).FirstOrDefault();

                System.Collections.Generic.IEnumerable<XElement> attrPos;
                if (table == null)
                {
                    foreach (var oeNode in node.Descendants(ns + "OE"))
                    {
                        if (oeNode.Descendants(ns + "T").Where(n => n.Attribute("selected") != null && n.Attribute("selected").Value == "all").Count() > 0
                                        && oeNode.Descendants(ns + "T").Where(n => n.Attribute("selected") == null || n.Attribute("selected").Value == "none").Count() > 0)
                        {
                            return true;
                        }
                    }
                }
                else
                {
                    foreach (var oeNode in table.Descendants(ns + "OE"))
                    {
                        var allSel = oeNode.Descendants(ns + "T").Any(n => n.Attribute("selected")?.Value == "all");
                        var unsel  = oeNode.Descendants(ns + "T").Any(n => n.Attribute("selected") == null || n.Attribute("selected").Value == "none");
                        if (allSel && unsel) return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// 產生 XML 插入至 OneNote
        /// Generate XML Insert To OneNote
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
                    //outline.Descendants().Where(n => n.Attribute("selected") != null && n.Attribute("selected").Value == "all").Remove();
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
                        int bcIndex = defaultStyle.IndexOf("background-color");
                        defaultStyle = defaultStyle.Remove(bcIndex, defaultStyle.IndexOf(';', bcIndex) - bcIndex +1);
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

                    //string nr = string.Format(@"<body style=""font-family:{0}"">", GenerateHighlightContent.GenerateHighLight.Config.OutputArguments["Font"].Value) +
                    //        itemNr.Replace("&apos;", "'") + "</body>";
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
                //string s = item.Replace(@"style=""", string.Format(@"style=""font-family:{0}; ", GenerateHighlightContent.GenerateHighLight.Config.OutputArguments["Font"].Value));
                //string s = string.Format(@"<body style=""font-family:{0}"">", GenerateHighlightContent.GenerateHighLight.Config.OutputArguments["Font"].Value) + 
                //            itemLine.Replace("&apos;", "'") + "</body>";
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
