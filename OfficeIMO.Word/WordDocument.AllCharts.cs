using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;

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
    /// Lists every chart drawing in the given content in document order, at any depth: inside tables, content controls,
    /// hyperlinks, fields, custom XML and tracked insertions. Deleted and moved-from content is skipped, and
    /// only the branch the Word reader selects from each markup-compatibility block is read.
    /// </summary>
    internal List<WordChart> CollectCharts(IEnumerable<OpenXmlElement> content) {
        var charts = new List<WordChart>();
        CollectCharts(content, charts);
        return charts;
    }

    private void CollectCharts(IEnumerable<OpenXmlElement> content, List<WordChart> charts) {
        foreach (OpenXmlElement element in content) {
            switch (element) {
                case DeletedRun:
                case MoveFromRun:
                    break;
                case AlternateContent alternateContent:
                    if (WordAlternateContentResolver.SelectBranch(alternateContent) is { } branch) {
                        CollectCharts(branch.ChildElements, charts);
                    }
                    break;
                case WordDrawing drawing when WordParagraph.DrawingOwnsChart(drawing):
                    Run? run = drawing.Ancestors<Run>().FirstOrDefault();
                    Paragraph? paragraph = run?.Ancestors<Paragraph>().FirstOrDefault();
                    if (run != null && paragraph != null) {
                        charts.Add(new WordChart(this, new WordParagraph(this, paragraph, run), drawing));
                    }
                    break;
                default:
                    if (element.HasChildren) CollectCharts(element.ChildElements, charts);
                    break;
            }
        }
    }
}
