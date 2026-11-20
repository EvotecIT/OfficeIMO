using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>
    /// Gets every chart that appears in the document: the body first, then the headers and footers each section
    /// defines, then footnotes and endnotes.
    /// </summary>
    /// <remarks>
    /// <see cref="Charts"/> keeps listing only the body, like <see cref="Images"/> and <see cref="Tables"/>.
    /// A header or footer shared by several sections is listed once.
    /// </remarks>
    public IReadOnlyList<WordChart> AllCharts {
        get {
            var charts = new List<WordChart>(Charts);
            var visited = new HashSet<OpenXmlElement>();
            using (UseBodySectionSnapshot()) {
                foreach (WordSection section in Sections) {
                    WordHeaderFooter?[] stories = {
                        section.Headers.Default, section.Headers.First, section.Headers.Even,
                        section.Footers.Default, section.Footers.First, section.Footers.Even
                    };
                    foreach (WordHeaderFooter? story in stories) {
                        OpenXmlCompositeElement? root = (OpenXmlCompositeElement?)story?._header ?? story?._footer;
                        if (root != null && visited.Add(root)) charts.AddRange(CollectCharts(root.ChildElements));
                    }
                }
            }

            var main = _wordprocessingDocument.MainDocumentPart;
            if (main?.FootnotesPart?.Footnotes is { } footnotes) {
                foreach (Footnote note in footnotes.Elements<Footnote>().Where(WordNoteResolver.IsOrdinary)) charts.AddRange(CollectCharts(note.ChildElements));
            }
            if (main?.EndnotesPart?.Endnotes is { } endnotes) {
                foreach (Endnote note in endnotes.Elements<Endnote>().Where(WordNoteResolver.IsOrdinary)) charts.AddRange(CollectCharts(note.ChildElements));
            }
            return charts.AsReadOnly();
        }
    }

    /// <summary>
    /// Lists the charts in block content in document order, including tables and block content controls, and reading
    /// only the selected branch of each markup-compatibility block.
    /// </summary>
    internal List<WordChart> CollectCharts(IEnumerable<OpenXmlElement> blockContent) {
        var paragraphs = new List<Paragraph>();
        CollectBlockParagraphs(blockContent, paragraphs);
        return WordSection.ConvertParagraphsToWordParagraphs(this, paragraphs)
            .SelectMany(paragraph => paragraph.EnumerateCharts())
            .ToList();
    }

    private static void CollectBlockParagraphs(IEnumerable<OpenXmlElement> elements, List<Paragraph> paragraphs) {
        foreach (OpenXmlElement element in elements) {
            switch (element) {
                case Paragraph paragraph:
                    paragraphs.Add(paragraph);
                    break;
                case AlternateContent alternateContent:
                    if (WordAlternateContentResolver.SelectBranch(alternateContent) is { } branch) {
                        CollectBlockParagraphs(branch.ChildElements, paragraphs);
                    }
                    break;
                case Table or TableRow or TableCell
                    or SdtBlock or SdtContentBlock or SdtRow or SdtContentRow or SdtCell or SdtContentCell
                    or CustomXmlBlock or CustomXmlRow or CustomXmlCell:
                    CollectBlockParagraphs(element.ChildElements, paragraphs);
                    break;
            }
        }
    }
}
