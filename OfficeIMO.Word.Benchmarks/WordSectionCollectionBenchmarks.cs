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
        WordBenchmarkCorpus.Check("OfficeIMO", string.Empty,
            $"SectionCount={SectionCount}; every document-wide list matches the sum of its sections and the written content",
            ValidateCollections);
    }

    private void ValidateCollections() {
        WordDocument document = _document ?? throw new InvalidOperationException("The benchmark document was not loaded.");
        EnsureEqual("Sections", SectionCount, document.Sections.Count);
        EnsureEqual("Section paragraphs", document.Sections.Sum(section => section.Paragraphs.Count), document.Paragraphs.Count);
        EnsureEqual("Tables", SectionCount * TablesPerSection, document.Tables.Count);
        EnsureEqual("Hyperlink paragraphs", SectionCount, document.ParagraphsHyperLinks.Count);
        EnsureEqual("Elements", document.Sections.Sum(section => section.Elements.Count), document.Elements.Count);
        EnsureEqual("Fields", SectionCount, document.Fields.Count);
        EnsureEqual("Hyperlinks", SectionCount, document.HyperLinks.Count);
        EnsureEqual("Check boxes", SectionCount, document.CheckBoxes.Count);
        EnsureEqual("Footnotes", SectionCount, document.FootNotes.Count);
        EnsureEqual("Nested tables included", SectionCount * TablesPerSection, document.TablesIncludingNestedTables.Count);
    }

    [GlobalCleanup]
    public void Cleanup() => _document?.Dispose();

    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:ParagraphsAndTables")]
    public int ParagraphsAndTables() => _document!.Paragraphs.Count + _document.Tables.Count;

    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:HyperlinkParagraphs")]
    public int HyperlinkParagraphs() => _document!.ParagraphsHyperLinks.Count;

    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:Elements")]
    public int Elements() => _document!.Elements.Count;

    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:FieldsHyperlinksAndBookmarks")]
    public int FieldsHyperlinksAndBookmarks() => _document!.Fields.Count + _document.HyperLinks.Count + _document.Bookmarks.Count;

    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:ControlsNotesAndNestedTables")]
    public int ControlsNotesAndNestedTables() =>
        _document!.CheckBoxes.Count + _document.StructuredDocumentTags.Count + _document.FootNotes.Count + _document.TablesIncludingNestedTables.Count;

    private static void EnsureEqual(string what, int expected, int actual) {
        if (actual != expected) {
            throw new InvalidOperationException(what + " returned " + actual + "; expected " + expected + ".");
        }
    }
}
