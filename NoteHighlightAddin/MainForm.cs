using GenerateHighlightContent;
using Helper;
using ICSharpCode.TextEditor.Document;
using NoteHighlightForm;
using NoteHighlightAddin.Preview;
using System;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace NoteHighlightAddin
{
    public partial class MainForm : Form
    {
        #region -- Field and Property --
        private const string span = "</span>";

        private string _codeType;

        private string _fileName;

        private HighLightParameter _parameters;
        private bool _darkMode;

        private string CodeContent { get { return this.txtCode.Text; } }

        private string CodeStyle { get { return this.cbx_style.Text; } }

        private bool IsShowLineNumber { get { return this.cbx_lineNumber.Checked; } }

        private bool IsClipboard { get { return this.cbx_Clipboard.Checked; } }

        private Color BackgroundColor => this.btnBackground.BackColor;

        public HighLightParameter Parameters => _parameters;

        private readonly bool _quickStyle;

        public bool DarkMode => this.cbx_darkMode.Checked;

        /// <summary>
        /// Set while the splitter is being positioned programmatically (initial load,
        /// resize-driven re-apply) so the SplitterMoved handler does not write the
        /// transient value back to user.config. Cleared once the user takes control.
        /// </summary>
        private bool _suppressSplitterPersist;

        /// <summary>
        /// False until MainForm_Shown has finished its startup sizing. The form opens
        /// proportional to the monitor work area, and with FixedPanel.None that resize
        /// makes the SplitContainer rescale the splitter to its designer ratio and fire
        /// SplitterMoved before the saved distance is applied. Persisting that transient
        /// value would overwrite the user's saved split (and resurrect the old default)
        /// on every launch, so SplitContainer_SplitterMoved ignores moves until the form
        /// is settled and only genuine user drags are saved thereafter.
        /// </summary>
        private bool _splitterReady;

        #endregion

        #region -- Constructor --

        public MainForm(string codeType, string fileName, string selectedText, bool quickStyle, bool darkMode)
        {
            _codeType = codeType;
            _fileName = fileName;
            InitializeComponent();
            LoadThemes();
            txtCode.Text = selectedText;
            _quickStyle = quickStyle;
            _darkMode = darkMode;

            if (_quickStyle)
            {
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
            }

            _previewPane.QuickStyleSuppress = _quickStyle;

            this.txtCode.TextChanged += (s, e) => SchedulePreview();
            this.cbx_style.SelectedIndexChanged += (s, e) => SchedulePreview();
            this.cbx_lineNumber.CheckedChanged += (s, e) => SchedulePreview();
            this.cbx_tableBorder.CheckedChanged += (s, e) => SchedulePreview();
            this.cbx_darkMode.CheckedChanged += (s, e) => SchedulePreview();
            this.btnBackground.BackColorChanged += (s, e) => SchedulePreview();

            this.splitContainer.SplitterMoved += SplitContainer_SplitterMoved;
        }

        /// <summary>
        /// Applies the persisted splitter percentage to the SplitContainer.
        /// Called after Shown so the SplitContainer has its final width and the
        /// Panel1/Panel2 MinSize clamps are honoured by the framework.
        /// </summary>
        private void ApplySavedSplitterDistance()
        {
            if (this.splitContainer == null) return;
            int width = this.splitContainer.Width;
            if (width <= 0) return;

            int percent = NoteHighlightForm.Properties.Settings.Default.MainFormPreviewSplitter;
            if (percent <= 0 || percent >= 100) percent = 50;

            int desired = (int)Math.Round(width * (percent / 100.0));

            // Honour Panel1MinSize and Panel2MinSize. The SplitContainer would throw
            // an InvalidOperationException if we set a value that violates either.
            int min = this.splitContainer.Panel1MinSize;
            int max = width - this.splitContainer.Panel2MinSize - this.splitContainer.SplitterWidth;
            if (max < min) return; // not enough room yet; leave designer default

            int clamped = Math.Max(min, Math.Min(max, desired));

            _suppressSplitterPersist = true;
            try
            {
                this.splitContainer.SplitterDistance = clamped;
            }
            finally
            {
                _suppressSplitterPersist = false;
            }
        }

        private void SplitContainer_SplitterMoved(object sender, SplitterEventArgs e)
        {
            if (_suppressSplitterPersist) return;
            if (!_splitterReady) return;
            if (!this.IsHandleCreated) return;

            int width = this.splitContainer.Width;
            if (width <= 0) return;

            int percent = (int)Math.Round(this.splitContainer.SplitterDistance * 100.0 / width);
            if (percent < 1) percent = 1;
            if (percent > 99) percent = 99;

            if (NoteHighlightForm.Properties.Settings.Default.MainFormPreviewSplitter == percent) return;

            NoteHighlightForm.Properties.Settings.Default.MainFormPreviewSplitter = percent;
            SettingsHelper.SafeSave();
        }

        private void SchedulePreview()
        {
            if (_quickStyle) return;
            if (_previewPane == null) return;

            var parameters = new HighLightParameter()
            {
                FileName = _fileName,
                Content = CodeContent,
                CodeType = _codeType,
                HighLightStyle = CodeStyle,
                ShowLineNumber = IsShowLineNumber,
                HighlightColor = BackgroundColor,
                ShowTableBorder = this.cbx_tableBorder.Checked,
                Font = NoteHighlightForm.Properties.Settings.Default.Font,
                FontSize = NoteHighlightForm.Properties.Settings.Default.FontSize
            };

            _previewPane.Render(parameters, this.DarkMode);
        }

        private void LoadThemes()
        {
            try
            {
                HighLightSection section = (new GenerateHighLight(AddIn.GetAddinDirectory())).Config;
                // Use typeof(MainForm).Assembly rather than GetCallingAssembly():
                // GetCallingAssembly is sensitive to JIT inlining and cross-AppDomain
                // COM callers, so it may resolve to mscorlib or ONENOTE.EXE instead of
                // the add-in assembly that actually ships the themes folder.
                var assemblyLocation = typeof(MainForm).Assembly.Location;
                if (string.IsNullOrEmpty(assemblyLocation))
                    assemblyLocation = new Uri(typeof(MainForm).Assembly.CodeBase).LocalPath;
                var workingDirectory = Path.Combine(ProcessHelper.GetDirectoryFromPath(assemblyLocation), section.FolderName, section.ThemeFolder);

                string[] files = Directory.GetFiles(workingDirectory, "*.theme");

                foreach (var item in files)
                {
                    cbx_style.Items.Add(Path.GetFileNameWithoutExtension(item));
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception from MainForm.LoadThemes:" + e.Message);
                return;
            }

            
            
        }

        #endregion

        #region -- Event --

        /// <summary>
        /// Form Load
        /// </summary>
        private void CodeForm_Load(object sender, EventArgs e)
        {
            this.txtCode.Document.HighlightingStrategy = HighlightingStrategyFactory.CreateHighlightingStrategy(CodeTypeTransform(_codeType));
            this.txtCode.Encoding = Encoding.UTF8;
            this.cbx_style.SelectedIndex = NoteHighlightForm.Properties.Settings.Default.HighLightStyle;
            this.btnBackground.BackColor = NoteHighlightForm.Properties.Settings.Default.BackgroundColor;
            this.cbx_Clipboard.Checked = NoteHighlightForm.Properties.Settings.Default.SaveOnClipboard;
            this.cbx_lineNumber.Checked = NoteHighlightForm.Properties.Settings.Default.ShowLineNumber;
            this.cbx_tableBorder.Checked = NoteHighlightForm.Properties.Settings.Default.ShowTableBorder;
            this.cbx_darkMode.Checked = _darkMode;
        }

        /// <summary>
        /// Form Closed
        /// </summary>
        private void CodeForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _previewPane?.Dispose();
            SaveSetting();
        }

        /// <summary>
        /// Generate HighLight File
        /// </summary>
        private void btnCodeHighLight_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(CodeStyle))
            {
                MessageBox.Show("Please select code Style!");
                return;
            }
            IGenerateHighLight generate = new GenerateHighLight(AddIn.GetAddinDirectory());

            string outputFileName = String.Empty;

            _parameters = new HighLightParameter()
            {
                FileName = _fileName,
                Content = CodeContent,
                CodeType = _codeType,
                HighLightStyle = CodeStyle,
                ShowLineNumber = IsShowLineNumber,
                HighlightColor = BackgroundColor,
                Font = NoteHighlightForm.Properties.Settings.Default.Font,
                FontSize = NoteHighlightForm.Properties.Settings.Default.FontSize

            };

            try
            {
                outputFileName = generate.GenerateHighLightCode(_parameters);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                this.Dispose();
                this.Close();
                return;
            }

            if (IsClipboard && !String.IsNullOrEmpty(outputFileName))
                InsertToClipboard(outputFileName);

            SaveSetting();

            this.Dispose();
            this.Close();
        }

        #endregion

        /// <summary>
        /// Copy HighLight Code To Clipboard
        /// </summary>
        private void InsertToClipboard(string outputFileName)
        {
            StringBuilder sb = new StringBuilder();

            using (FileStream fs = new FileStream(outputFileName, FileMode.Open, FileAccess.Read))
            {
                using (StreamReader sr = new StreamReader(fs, new UTF8Encoding(false)))
                {
                    while (sr.Peek() >= 0)
                    {
                        string line = sr.ReadLine();

                        string byteOrderMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());
                        line = line.Replace(byteOrderMarkUtf8, "");

                        if (line.StartsWith("<pre") && this.DarkMode)
                        {

                            //Remove background-color element so that text would render with correct contrast in dark mode
                            line = PreviewHtmlWrapper.StripPreBackgroundColor(line);
                        }


                        if (!line.StartsWith("</pre>"))
                        {
                            line = line.Replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;").Replace("&apos;", "'") + "<br />";
                        }
                        var charList = line.ToCharArray().ToList();

                        StringBuilder sbLine = new StringBuilder();
                        int index = 0;

                        if (IsShowLineNumber && !line.StartsWith("</pre>"))
                        {
                            index = line.IndexOf(span) + span.Length;
                            string nrLine = line.Substring(0, index);

                            int endTextIndex = nrLine.IndexOf(span);
                            int startTextIndex = nrLine.LastIndexOf(">", endTextIndex) + 1;

                            nrLine = nrLine.Substring(0, startTextIndex) + nrLine.Substring(startTextIndex, endTextIndex-startTextIndex).Replace(" ", "&nbsp;") + nrLine.Substring(endTextIndex);

                            sbLine.Append(nrLine);
                        }

                        for (int i = index; i < charList.Count; i++)
                        {
                            if (charList[i] == ' ')
                            {
                                sbLine.Append("&nbsp;&nbsp;");
                            }
                            else
                            {
                                sbLine.Append(line.Substring(i));
                                break;
                            }
                        }
                        sb.AppendLine(sbLine.ToString());
                    }
                }
            }
            HtmlFragment.CopyToClipboard(sb.ToString());
            File.Delete(outputFileName);
        }

        /// <summary>
        /// Transfer ICSharpCode.TextEditor Control Use FileCode
        /// </summary>
        private string CodeTypeTransform(string codeType)
        {
            string result = string.Empty;
            switch (codeType.ToLower())
            {
                case "cs":
                    result = "C#";
                    break;
                case "vb":
                    result = "VBNET";
                    break;
                case "js":
                    result = "JavaScript";
                    break;
                case "xml":
                    result = "XML";
                    break;
                case "css":
                    result = "CSS";
                    break;
                case "html":
                    result = "HTML";
                    break;
                case "php":
                    result = "PHP";
                    break;
                case "java":
                    result = "Java";
                    break;
                case "c":
                    result = "C++.NET";
                    break;
                default:
                    result = "";
                    break;
            }
            return result;
        }

        /// <summary>
        /// Save User Setting
        /// </summary>
        private void SaveSetting()
        {
            var defaultSettings = NoteHighlightForm.Properties.Settings.Default;
            defaultSettings.ShowLineNumber = this.cbx_lineNumber.Checked;
            defaultSettings.SaveOnClipboard = this.cbx_Clipboard.Checked;
            defaultSettings.ShowTableBorder = this.cbx_tableBorder.Checked;
            defaultSettings.DarkMode = this.cbx_darkMode.Checked;
            defaultSettings.HighLightStyle = this.cbx_style.SelectedIndex;
            defaultSettings.BackgroundColor = this.btnBackground.BackColor;
            SettingsHelper.SafeSave();
        }

        private void btnBackground_Click(object sender, EventArgs e)
        {
            contextMenuStrip1.Show(btnBackground, new Point(0, btnBackground.Height));
            
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {

            if (_quickStyle)
            {
                btnCodeHighLight.PerformClick();
            }
            else
            {
                // This is necessary in order for SetForegroundWindow to work consistently
                this.WindowState = FormWindowState.Minimized;
                this.WindowState = FormWindowState.Normal;

                NativeMethods.SetForegroundWindow(this.Handle);

                // Open proportionally to the active monitor's work area (the screen the
                // cursor is on), clamped to MinimumSize so the form never collapses on a
                // small/low-res display. CenterScreen does not re-centre after a post-Shown
                // resize, so place the form manually at the centre of that work area.
                Rectangle workArea = Screen.FromPoint(Cursor.Position).WorkingArea;
                int targetWidth = Math.Max(this.MinimumSize.Width, (int)(workArea.Width * 0.70));
                int targetHeight = Math.Max(this.MinimumSize.Height, (int)(workArea.Height * 0.75));
                this.Size = new Size(targetWidth, targetHeight);
                this.Location = new Point(
                    workArea.Left + (workArea.Width - this.Width) / 2,
                    workArea.Top + (workArea.Height - this.Height) / 2);

                ApplySavedPreviewVisibility();
                if (!this.splitContainer.Panel2Collapsed)
                {
                    ApplySavedSplitterDistance();
                }

                this.BeginInvoke(new Action(SchedulePreview));

                // Arm splitter persistence only after the startup resize/layout has
                // settled. Queued behind SchedulePreview so any resize-induced
                // SplitterMoved from the proportional sizing above is ignored; from
                // here on only genuine user drags are written to user.config.
                this.BeginInvoke(new Action(() => _splitterReady = true));
            }

        }

        /// <summary>
        /// Reads the persisted preview-visibility flag and applies it to the SplitContainer.
        /// Must be called before ApplySavedSplitterDistance: when Panel2Collapsed is true the
        /// splitter distance is invalid (the framework will throw).
        /// </summary>
        private void ApplySavedPreviewVisibility()
        {
            bool visible = NoteHighlightForm.Properties.Settings.Default.MainFormPreviewVisible;
            this.splitContainer.Panel2Collapsed = !visible;
            UpdateTogglePreviewButtonText(visible);
        }

        private void UpdateTogglePreviewButtonText(bool previewVisible)
        {
            this.btnTogglePreview.Text = previewVisible ? "Hide preview" : "Show preview";
        }

        private void btnTogglePreview_Click(object sender, EventArgs e)
        {
            // Toggle the persisted flag, then drive Panel2Collapsed off the new value.
            bool nowVisible = this.splitContainer.Panel2Collapsed; // collapsed -> will become visible
            this.splitContainer.Panel2Collapsed = !nowVisible;
            UpdateTogglePreviewButtonText(nowVisible);

            NoteHighlightForm.Properties.Settings.Default.MainFormPreviewVisible = nowVisible;
            SettingsHelper.SafeSave();

            if (nowVisible)
            {
                // Re-showing: reapply the saved splitter distance and refresh the preview
                // because the pane was not rendering while hidden.
                ApplySavedSplitterDistance();
                SchedulePreview();
            }
        }

        private void PickColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                btnBackground.BackColor = colorDialog1.Color;
            }
        }

        private void TransparentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            btnBackground.BackColor = Color.Transparent;
        }
    }
}
