using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    private static readonly char[] XmlTextEdgeWhitespace = [' ', '\t', '\r', '\n'];

    /// <summary>Reads significant WordprocessingML text while retaining non-breaking and other Unicode spaces.</summary>
    private static string ReadWordprocessingText(Text text) =>
        text.Space?.Value == SpaceProcessingModeValues.Preserve
            ? text.Text
            : text.Text.Trim(XmlTextEdgeWhitespace);
}
