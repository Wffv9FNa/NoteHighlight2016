using System.Drawing;

namespace GenerateHighlightContent
{
    public interface IGenerateHighLight
    {
        /// <summary> Produce highlighted code. </summary>
        /// <returns> Path of the generated output file. </returns>
        string GenerateHighLightCode(HighLightParameter parameter);
    }

    public class HighLightParameter
    {
        /// <summary> Content. </summary>
        public string Content { get; set; }

        /// <summary> Syntax/language. </summary>
        public string CodeType { get; set; }

        /// <summary> Highlight theme. </summary>
        public string HighLightStyle { get; set; }

        /// <summary> Whether to show line numbers. </summary>
        public bool ShowLineNumber { get; set; }

        /// <summary> File name. </summary>
        public string FileName { get; set; }

        public Color HighlightColor { get; set; }

        public string Font { get; set; }

        public int FontSize { get; set; }
    }
}
