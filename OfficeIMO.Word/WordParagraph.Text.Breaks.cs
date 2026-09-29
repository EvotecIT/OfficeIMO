using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace OfficeIMO.Word;

public partial class WordParagraph {
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
