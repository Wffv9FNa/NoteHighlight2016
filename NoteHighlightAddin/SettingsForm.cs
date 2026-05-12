using GenerateHighlightContent;
using Helper;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace NoteHighlightAddin
{
    public partial class SettingsForm : Form
    {
        private const string SampleSnippet = @"// Computes the n-th Fibonacci number using memoisation.
// Demonstrates: keywords, strings, numerics, generics, comments,
// one long line that should wrap or scroll.
using System;
using System.Collections.Generic;

public static class FibDemo
{
    private static readonly Dictionary<int, long> Cache = new Dictionary<int, long>();

    public static long Fib(int n)
    {
        if (n < 0) throw new ArgumentOutOfRangeException(""n"", ""Fibonacci is undefined for negative indices."");
        if (n < 2) return n;
        if (Cache.TryGetValue(n, out long cached)) return cached;
        long value = Fib(n - 1) + Fib(n - 2);
        Cache[n] = value;
        return value;
    }

    public static void Main() { Console.WriteLine(""Fib(40) = "" + Fib(40)); }
}";

        public SettingsForm()
        {
            InitializeComponent();

            fontDialog1.Font = new Font(NoteHighlightForm.Properties.Settings.Default.Font, NoteHighlightForm.Properties.Settings.Default.FontSize);
            btnFont.Text = "Font:" + fontDialog1.Font.Name + ", Size:" + fontDialog1.Font.Size;
            btnFont.Font = fontDialog1.Font;
            cbShowTableBorder.Checked = NoteHighlightForm.Properties.Settings.Default.ShowTableBorder;
        }

        private void BtnFont_Click(object sender, EventArgs e)
        {
            fontDialog1.Font = new Font(NoteHighlightForm.Properties.Settings.Default.Font, NoteHighlightForm.Properties.Settings.Default.FontSize);
            if (fontDialog1.ShowDialog() == DialogResult.OK)
            {
                btnFont.Text = "Font:"+fontDialog1.Font.Name + ", Size:" + fontDialog1.Font.Size;
                btnFont.Font = fontDialog1.Font;

                NoteHighlightForm.Properties.Settings.Default.Font = fontDialog1.Font.Name;
                NoteHighlightForm.Properties.Settings.Default.FontSize = (int)Math.Round(fontDialog1.Font.Size);

                SettingsHelper.SafeSave();
                SchedulePreview();
            }


        }

        private void ChShowTableBorder_CheckedChanged(object sender, EventArgs e)
        {
            NoteHighlightForm.Properties.Settings.Default.ShowTableBorder = cbShowTableBorder.Checked;

            SettingsHelper.SafeSave();
            SchedulePreview();
        }

        private void SchedulePreview()
        {
            if (_previewPane == null) return;

            var defaults = NoteHighlightForm.Properties.Settings.Default;
            var parameters = new HighLightParameter()
            {
                FileName = "settings-preview",
                Content = SampleSnippet,
                CodeType = "cs",
                HighLightStyle = ResolveThemeName(defaults.HighLightStyle),
                ShowLineNumber = defaults.ShowLineNumber,
                HighlightColor = defaults.BackgroundColor,
                Font = fontDialog1.Font.Name,
                FontSize = (int)Math.Round(fontDialog1.Font.Size)
            };

            _previewPane.Render(parameters, defaults.DarkMode);
        }

        private static string ResolveThemeName(int themeIndex)
        {
            try
            {
                var section = (new GenerateHighLight(AddIn.GetAddinDirectory())).Config;
                var assemblyLocation = typeof(SettingsForm).Assembly.Location;
                if (string.IsNullOrEmpty(assemblyLocation))
                    assemblyLocation = new Uri(typeof(SettingsForm).Assembly.CodeBase).LocalPath;
                var themesDir = Path.Combine(ProcessHelper.GetDirectoryFromPath(assemblyLocation), section.FolderName, section.ThemeFolder);
                var themes = Directory.GetFiles(themesDir, "*.theme");
                if (themes.Length == 0) return string.Empty;
                int idx = themeIndex >= 0 && themeIndex < themes.Length ? themeIndex : 0;
                return Path.GetFileNameWithoutExtension(themes[idx]);
            }
            catch
            {
                return string.Empty;
            }
        }

        private void SettingsForm_Shown(object sender, EventArgs e)
        {
            // This is necessary in order for SetForegroundWindow to work consistently
            this.WindowState = FormWindowState.Minimized;
            this.WindowState = FormWindowState.Normal;

            NativeMethods.SetForegroundWindow(this.Handle);

            this.BeginInvoke(new Action(SchedulePreview));
        }

        private void SettingsForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _previewPane?.Dispose();
        }
    }
}
