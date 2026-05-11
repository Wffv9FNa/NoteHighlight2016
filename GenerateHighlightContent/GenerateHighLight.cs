using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Configuration;
using Helper;

namespace GenerateHighlightContent
{
    public class GenerateHighLight : IGenerateHighLight
    {
        #region -- Field and Property --

        public string Content { get; set; }

        public string CodeType { get; set; }

        public string HighLightStyle { get; set; }

        public bool ShowLineNumber { get; set; }

        public string FileName { get; set; }

        public string Font { get; set; }

        public int FontSize { get; set; }

        /// <summary> highlight.exe 參數 設定於 App.config 的 HighLightSection 區塊 </summary>
        private HighLightSection _section;

        public HighLightSection Config { get { return _section; } }
        #endregion

        #region -- IGenerageHighLight Member --

        public GenerateHighLight()
        {
            Configuration c = ConfigurationManager.OpenExeConfiguration(Assembly.GetCallingAssembly().Location);
            _section = c.GetSection("HighLightSection") as HighLightSection;
        }

        /// <summary> 呼叫highlight.exe 產生高亮後的html </summary>
        /// <returns>回傳 Html 所在的路徑</returns>
        public string GenerateHighLightCode(HighLightParameter parameter)
        {
            InitParameter(parameter);

            string tempPath = Path.GetTempPath();
            string inputFileName = Path.Combine(tempPath, FileName);
            string outputFileName = Path.Combine(tempPath, FileName) + ".html";

            if (_section == null)
                throw new FileNotFoundException("ConfigurationManager.GetSection(\"HighLightSection\") failed!");

            var workingDirectory = Path.Combine(ProcessHelper.GetDirectoryFromPath(Assembly.GetCallingAssembly().Location), _section.FolderName);

            // Reject any value that could break out of the quoted argument or inject a new
            // highlight.exe switch (e.g. --plug-in=evil.lua). highlight.exe supports Lua
            // plug-ins which can run arbitrary host commands, so unchecked args = user-level RCE.
            ValidateArguments(workingDirectory);

            File.WriteAllText(inputFileName, Content, Encoding.UTF8);
            try
            {
                ProcessHelper helper = new ProcessHelper(workingDirectory, _section.ProcessName);
                helper.Arguments = GenerateArguments(inputFileName, outputFileName);
                helper.IsWaitForInputIdle = false;
                helper.WindowStyle = ProcessWindowStyle.Hidden;

                helper.ProcessStart();

                if (!File.Exists(outputFileName))
                    throw new FileNotFoundException("Can not find outputFile.");

                return outputFileName;
            }
            finally
            {
                // Always remove the input scratch file, even when highlight.exe fails to
                // produce output. Otherwise the user's raw source code is left in %TEMP%
                // indefinitely on every error path.
                try { if (File.Exists(inputFileName)) File.Delete(inputFileName); } catch { }
            }
        }

        /// <summary>
        /// Identifier-shaped values (codeType, theme name): letters, digits, '_', '+', '-', '.'.
        /// Length capped to keep error messages bounded.
        /// </summary>
        private static readonly Regex IdentifierRegex =
            new Regex(@"^[A-Za-z0-9_+.\-]{1,64}$", RegexOptions.Compiled);

        /// <summary>
        /// Reject font/codeType/theme values that could close the quoted token in
        /// GenerateArguments and append attacker-controlled highlight.exe switches.
        /// Allow-list HighLightStyle against the bundled themes folder.
        /// </summary>
        private void ValidateArguments(string workingDirectory)
        {
            ValidateFontName(Font);
            ValidateIdentifier("CodeType", CodeType);
            ValidateIdentifier("HighLightStyle", HighLightStyle);

            // FontSize is an int; nothing to do.

            string themesDir = Path.Combine(workingDirectory, _section.ThemeFolder);
            string themePath = Path.Combine(themesDir, HighLightStyle + ".theme");
            if (!File.Exists(themePath))
                throw new ArgumentException(
                    String.Format("Unknown highlight theme '{0}'.", HighLightStyle));
        }

        private static void ValidateIdentifier(string name, string value)
        {
            if (String.IsNullOrEmpty(value))
                throw new ArgumentException(name + " must not be empty.");
            if (value.StartsWith("-"))
                throw new ArgumentException(name + " must not start with '-'.");
            if (!IdentifierRegex.IsMatch(value))
                throw new ArgumentException(
                    name + " contains invalid characters (allowed: letters, digits, '_', '+', '-', '.').");
        }

        private static void ValidateFontName(string value)
        {
            if (String.IsNullOrEmpty(value))
                throw new ArgumentException("Font must not be empty.");
            if (value.Length > 128)
                throw new ArgumentException("Font name is too long.");
            if (value.IndexOf('"') >= 0)
                throw new ArgumentException("Font must not contain '\"'.");
            if (value.StartsWith("-"))
                throw new ArgumentException("Font must not start with '-'.");
            foreach (char c in value)
            {
                if (Char.IsControl(c))
                    throw new ArgumentException("Font must not contain control characters.");
            }
        }

        /// <summary> 初始化參數 </summary>
        private void InitParameter(HighLightParameter parameter)
        {
            Content = parameter.Content;
            CodeType = parameter.CodeType;
            HighLightStyle = parameter.HighLightStyle;
            ShowLineNumber = parameter.ShowLineNumber;
            FileName = parameter.FileName;
            Font = parameter.Font;
            FontSize = parameter.FontSize;
        }

        /// <summary> 產生HighLight.exe 所需的參數 </summary>
        private string GenerateArguments(string inputFileName, string outputFileName)
        {
            StringBuilder sb = new StringBuilder();

            ReadConfigCollection(sb, _section.GeneralArguments);
            ReadConfigCollection(sb, _section.OutputArguments);

            if (ShowLineNumber)
                sb.Append(" " + _section.OutputArguments["LineNumbers"].Key);

            string arguments = sb.ToString().TemplateSubstitute(new
            {
                inputFileName = String.Format("\"{0}\"", inputFileName),
                outputFileName = String.Format("\"{0}\"", outputFileName),
                codeType = CodeType,
                highLightStyle = HighLightStyle,
                font = String.Format("\"{0}\"", Font),
                fontSize = FontSize
            });

            return arguments;
        }

        /// <summary> 讀取 ConfigurationElementCollection </summary>
        private void ReadConfigCollection(StringBuilder sb, ConfigurationElementCollection collection)
        {
            foreach (Argument item in collection)
            {
                if (item.Option)
                    continue;

                sb.Append(item.Key);
                if (!String.IsNullOrEmpty(item.Value))
                    sb.Append(" " + (item.Value.Contains(" ") ? String.Format("\"{0}\"", item.Value) : item.Value));
                sb.Append(" ");
            }
        }

        #endregion
    }
}
