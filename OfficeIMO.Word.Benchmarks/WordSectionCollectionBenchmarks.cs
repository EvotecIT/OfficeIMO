using System.Globalization;
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>
/// Reads the document-wide lists of a multi-section document at two section counts, so the growth
/// between the sizes shows whether the lists scale with the document or with sections times the document.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "Read")]
public class WordSectionCollectionBenchmarks {
    private const int ParagraphsPerSection = 20;
    private const int TablesPerSection = 2;
    private WordDocument? _document;

    [Params(25, 100)]
    public int SectionCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create()) {
            for (int section = 0; section < SectionCount; section++) {
                WordSection target = section == 0 ? document.Sections[0] : document.AddSection();
                for (int paragraph = 0; paragraph < ParagraphsPerSection; paragraph++) {
                    target.AddParagraph("Section " + section.ToString(CultureInfo.InvariantCulture) + " paragraph " + paragraph.ToString(CultureInfo.InvariantCulture));
                }
                for (int table = 0; table < TablesPerSection; table++) {
                    target.AddTable(2, 2);
                }
                target.AddParagraph().AddHyperLink("Link", new Uri("https://example.com/"));
                WordParagraph extras = target.AddParagraph("Page ");
                extras.AddField(WordFieldType.Page);
                extras.AddBookmark("Bookmark" + section.ToString(CultureInfo.InvariantCulture));
                extras.AddCheckBox();
                extras.AddFootNote("Note " + section.ToString(CultureInfo.InvariantCulture));
            }
            bytes = document.ToBytes();
        }

        _document = WordDocument.Load(new MemoryStream(bytes, writable: false));
        EnsureEqual("Sections", SectionCount, _document.Sections.Count);
        EnsureEqual("Section paragraphs", _document.Sections.Sum(section => section.Paragraphs.Count), _document.Paragraphs.Count);
        EnsureEqual("Tables", SectionCount * TablesPerSection, _document.Tables.Count);
        EnsureEqual("Hyperlink paragraphs", SectionCount, _document.ParagraphsHyperLinks.Count);
        EnsureEqual("Elements", _document.Sections.Sum(section => section.Elements.Count), _document.Elements.Count);
        EnsureEqual("Fields", SectionCount, _document.Fields.Count);
        EnsureEqual("Hyperlinks", SectionCount, _document.HyperLinks.Count);
        EnsureEqual("Check boxes", SectionCount, _document.CheckBoxes.Count);
        EnsureEqual("Footnotes", SectionCount, _document.FootNotes.Count);
        EnsureEqual("Nested tables included", SectionCount * TablesPerSection, _document.TablesIncludingNestedTables.Count);
    }

    [GlobalCleanup]
    public void Cleanup() => _document?.Dispose();

    [Benchmark]
    public int ParagraphsAndTables() => _document!.Paragraphs.Count + _document.Tables.Count;

    [Benchmark]
    public int HyperlinkParagraphs() => _document!.ParagraphsHyperLinks.Count;

    [Benchmark]
    public int Elements() => _document!.Elements.Count;

    [Benchmark]
    public int FieldsHyperlinksAndBookmarks() => _document!.Fields.Count + _document.HyperLinks.Count + _document.Bookmarks.Count;

    [Benchmark]
    public int ControlsNotesAndNestedTables() =>
        _document!.CheckBoxes.Count + _document.StructuredDocumentTags.Count + _document.FootNotes.Count + _document.TablesIncludingNestedTables.Count;

    private static void EnsureEqual(string what, int expected, int actual) {
        if (actual != expected) {
            throw new InvalidOperationException(what + " returned " + actual + "; expected " + expected + ".");
        }
    }
}
