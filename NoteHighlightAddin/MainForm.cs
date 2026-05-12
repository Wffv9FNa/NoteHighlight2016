using GenerateHighlightContent;
using Helper;
using ICSharpCode.TextEditor.Document;
using NoteHighLightForm;
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

        public bool DarkMode => _darkMode;

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
        }

        /// <summary>
        /// Form Closed
        /// </summary>
        private void CodeForm_FormClosed(object sender, FormClosedEventArgs e)
        {
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
