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
                        if (root != null && visited.Add(root)) AddCharts(charts, root);
                    }
                }
            }

            var main = _wordprocessingDocument.MainDocumentPart;
            if (main?.FootnotesPart?.Footnotes is { } footnotes) {
                foreach (Footnote note in footnotes.Elements<Footnote>().Where(note => note.Type == null)) AddCharts(charts, note);
            }
            if (main?.EndnotesPart?.Endnotes is { } endnotes) {
                foreach (Endnote note in endnotes.Elements<Endnote>().Where(note => note.Type == null)) AddCharts(charts, note);
            }
            return charts.AsReadOnly();
        }
    }

    private void AddCharts(List<WordChart> charts, OpenXmlCompositeElement root) {
        foreach (WordParagraph paragraph in WordSection.ConvertParagraphsToWordParagraphs(this, root.Descendants<Paragraph>())) {
            if (paragraph.IsChart && paragraph.Chart is { } chart) charts.Add(chart);
        }
    }
}
