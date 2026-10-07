using OfficeIMO.Markdown;
using OfficeIMO.Reader;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderWordListMarkdownTests {
    [Theory]
    [InlineData(WordListLevelKind.BulletSolidRound, "•", "-", false)]
    [InlineData(WordListLevelKind.DecimalDot, "12.", "12.", true)]
    [InlineData(WordListLevelKind.LowerRomanDot, "xii.", "12.", true)]
    [InlineData(WordListLevelKind.UpperLetterDot, "L.", "12.", true)]
    public void WordPageMarkdownUsesPortableListSyntaxAndRetainsTheAuthoredMarker(WordListLevelKind kind, string marker, string portable, bool ordered) {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            WordList list = document.AddCustomList();
            list.Numbering.AddLevel(new WordListLevel(kind).SetStartNumberingValue(12));
            list.AddItem("First item");
            document.Save();
        }
        stream.Position = 0;
        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.Word().ReadDocument(stream, "list.docx");
        OfficeDocumentBlock block = Assert.Single(result.Blocks, item => item.Kind == "list-item");
        var page = new OfficeDocumentReadResult { Pages = new[] { new OfficeDocumentPage { Blocks = new[] { block } } } };
        string markdown = Assert.Single(page.GetPageMarkdown(new OfficeDocumentPageMarkdownOptions { IncludePageMarkers = false })).Markdown;
        Assert.Equal(portable + " First item", markdown);
        Assert.Equal(marker, block.Marker);
        Assert.Equal(ordered ? 12 : (int?)null, block.ListIndex);
        OfficeDocumentReadResult transported = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(page));
        Assert.Equal(markdown, Assert.Single(transported.GetPageMarkdown(new OfficeDocumentPageMarkdownOptions { IncludePageMarkers = false })).Markdown);
        IMarkdownBlock parsed = Assert.Single(MarkdownReader.Parse(markdown).Blocks);
        if (ordered) Assert.Equal(12, Assert.IsType<OrderedListBlock>(parsed).Start);
        else Assert.IsType<UnorderedListBlock>(parsed);
    }
}
