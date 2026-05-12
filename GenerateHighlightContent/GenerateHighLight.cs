using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Configuration;
using System.Threading;
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

        /// <summary> highlight.exe arguments configured in the HighLightSection block of App.config. </summary>
        private HighLightSection _section;

        /// <summary>
        /// The add-in install directory. Captured at construction time so the
        /// highlight.exe working directory is derived from the same authoritative
        /// path the caller used to resolve the config file - not re-probed via
        /// Assembly.Location, which is unreliable under COM activation.
        /// </summary>
        private readonly string _addinDirectory;

        public HighLightSection Config { get { return _section; } }
        #endregion

        #region -- IGenerageHighLight Member --

        /// <summary>
        /// Production constructor. The caller (the NoteHighlightAddin DLL) knows
        /// its own install location authoritatively and passes it in. The
        /// HighLightSection lives in NoteHighlightAddin.dll.config alongside
        /// NoteHighlightAddin.dll; the GenerateHighlightContent App.config is
        /// not copied to output and must not be relied on.
        /// </summary>
        /// <param name="addinDirectory">Directory containing NoteHighlightAddin.dll
        /// and its .config file. Must not be null or empty.</param>
        public GenerateHighLight(string addinDirectory)
        {
            if (string.IsNullOrEmpty(addinDirectory))
                throw new ArgumentException(
                    "addinDirectory must not be null or empty.", "addinDirectory");

            _addinDirectory = addinDirectory;

            var configPath = Path.Combine(addinDirectory, "NoteHighlightAddin.dll.config");
            if (!File.Exists(configPath))
            {
                throw new FileNotFoundException(
                    "NoteHighlightAddin.dll.config not found at expected install path '" +
                    configPath + "'.", configPath);
            }

            var map = new ExeConfigurationFileMap { ExeConfigFilename = configPath };
            Configuration c = ConfigurationManager.OpenMappedExeConfiguration(
                map, ConfigurationUserLevel.None);
            _section = c.GetSection("HighLightSection") as HighLightSection;
            if (_section == null)
            {
                throw new ConfigurationErrorsException(
                    "HighLightSection not found. Tried config path: '" + configPath +
                    "' (File.Exists=" + File.Exists(configPath) + ").");
            }

            Trace.TraceInformation(
                "NoteHighlight2016: loaded HighLightSection from '" + configPath + "'.");
        }

        /// <summary>
        /// Compatibility constructor for COM activation and MSTest. Production
        /// callers MUST use the constructor that takes an explicit add-in
        /// directory - this overload depends on Assembly.Location, which is
        /// unreliable under cross-AppDomain COM activation (it can return the
        /// empty string for assemblies loaded from byte arrays or shadow copies).
        /// </summary>
        public GenerateHighLight() : this(BestEffortAddinDirectory())
        {
        }

        private static string BestEffortAddinDirectory()
        {
            var asm = typeof(GenerateHighLight).Assembly;
            string path = null;
            try
            {
                path = asm.Location;
            }
            catch (NotSupportedException) { /* dynamic assembly */ }

            if (string.IsNullOrEmpty(path))
            {
                try
                {
                    var cb = asm.CodeBase;
                    if (!string.IsNullOrEmpty(cb))
                        path = new Uri(cb).LocalPath;
                }
                catch (NotSupportedException) { /* dynamic assembly */ }
                catch (UriFormatException) { /* malformed CodeBase URI */ }
            }

            if (string.IsNullOrEmpty(path))
                throw new InvalidOperationException(
                    "Could not determine the add-in install directory. Use the " +
                    "GenerateHighLight(string addinDirectory) constructor and pass " +
                    "the path explicitly.");

            return Path.GetDirectoryName(path);
        }

        /// <summary> Invoke highlight.exe to produce the highlighted HTML. </summary>
        /// <returns>The path to the generated HTML file.</returns>
        public string GenerateHighLightCode(HighLightParameter parameter)
        {
            return GenerateHighLightCode(parameter, CancellationToken.None);
        }

        /// <summary>
        /// Cancellable variant of <see cref="GenerateHighLightCode(HighLightParameter)"/>.
        /// Threads the supplied <paramref name="cancellationToken"/> into
        /// <see cref="ProcessHelper.ProcessStart(CancellationToken)"/> so a
        /// preview render in flight can be terminated when a fresher edit
        /// arrives. The input scratch file is still deleted on every path via
        /// the <c>finally</c> block, including the cancellation path.
        /// </summary>
        public string GenerateHighLightCode(HighLightParameter parameter, CancellationToken cancellationToken)
        {
            InitParameter(parameter);

            string tempPath = Path.GetTempPath();
            string inputFileName = Path.Combine(tempPath, FileName);
            string outputFileName = Path.Combine(tempPath, FileName) + ".html";

            if (_section == null)
                throw new FileNotFoundException("ConfigurationManager.GetSection(\"HighLightSection\") failed!");

            // Use the add-in directory the caller supplied at construction time.
            // Do NOT re-derive this from Assembly.Location: that path can return
            // string.Empty under cross-AppDomain COM activation, and we would
            // then resolve highlight.exe against OneNote's CWD (Office16).
            var workingDirectory = Path.Combine(_addinDirectory, _section.FolderName);

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

                helper.ProcessStart(cancellationToken);

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

        /// <summary> Initialise parameters. </summary>
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

        /// <summary> Build the argument string passed to highlight.exe. </summary>
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

        /// <summary> Read a ConfigurationElementCollection into the argument buffer. </summary>
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
