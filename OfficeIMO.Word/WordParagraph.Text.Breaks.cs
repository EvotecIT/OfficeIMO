using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>Projects actual page and column breaks as text separators without changing literal U+2028 text.</summary>
    internal string GetTextWithNonTextBreakReplacement(char replacement) {
        string text = Text;
        IReadOnlyDictionary<int, WordBreakType>? breaks = GetNonTextBreakPositions();
        if (breaks == null) return text;
        char[] projected = text.ToCharArray();
        foreach (int position in breaks.Keys) projected[position] = replacement;
        return new string(projected);
    }

    /// <summary>Reads visible XML text with a selected separator only at actual page and column break positions.</summary>
    internal static string ReadVisibleText(OpenXmlElement element, char nonTextBreakReplacement) {
        var text = new StringBuilder();
        var breaks = new Dictionary<int, WordBreakType>();
        AppendVisibleText(text, element, breaks);
        foreach (int position in breaks.Keys) text[position] = nonTextBreakReplacement;
        return text.ToString();
    }

    // Positions share the visible-text traversal, so literal U+2028 text is never mistaken for a page break.
    internal IReadOnlyDictionary<int, WordBreakType>? GetNonTextBreakPositions() {
        if (_run == null && _stdRun?.SdtProperties?.Elements<DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox>().Any() == true)
            return null;
        OpenXmlElement? element = _visibleRun ?? _run ?? (OpenXmlElement?)_stdRun ?? _hyperlink ?? (OpenXmlElement?)_simpleField;
        if (element == null && _runs != null) {
            if (!_runs.Any(run => run.Descendants<Break>().Any(node => !IsTextWrappingBreak(node)))) return null;
            var fields = new Dictionary<int, WordBreakType>();
            ReadComplexFieldResultText(_runs, fields);
            return fields.Count == 0 ? null : fields;
        }
        element ??= (OpenXmlElement?)_stdRun ?? _officeMath ?? (OpenXmlElement?)_mathParagraph;
        if (element == null || !element.Descendants<Break>().Any(node => !IsTextWrappingBreak(node))) return null;
        var breaks = new Dictionary<int, WordBreakType>();
        AppendVisibleText(new StringBuilder(), element, breaks);
        return breaks.Count == 0 ? null : breaks;
    }
}
