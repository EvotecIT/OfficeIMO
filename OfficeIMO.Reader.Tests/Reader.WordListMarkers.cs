using OfficeIMO.Reader;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderWordListMarkerTests {
    [Fact]
    public void WordReaderRejectsUnboundedGeneratedListMarkers() {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            WordList list = document.AddCustomList();
            list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.UpperRomanDot));
            list.Numbering.Levels[0].StartNumberingValue = int.MaxValue;
            list.AddItem("Extreme Roman item");
            document.Save();
        }
        stream.Position = 0;
        Assert.Contains("list marker", Assert.Throws<InvalidDataException>(() =>
            OfficeIMO.Reader.Tests.ReaderTestReaders.Word().ReadDocument(stream, "bounded.docx")).Message);
    }

    [Theory]
    [InlineData(WordListLevelKind.BulletSolidRound, "•")]
    [InlineData(WordListLevelKind.None, "")]
    public void WordRichBlocksRetainPortableBulletsAndInvisibleMarkers(WordListLevelKind kind, string marker) {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            WordList list = document.AddCustomList();
            list.Numbering.AddLevel(new WordListLevel(kind));
            list.AddItem("List item");
            document.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.Word().ReadDocument(stream, "bulleted.docx");
        Assert.Equal(marker, Assert.Single(result.Blocks.Where(block => block.Kind == "list-item")).Marker);
    }

    [Theory]
    [InlineData(WordListLevelKind.DecimalDot, "12.", "13.")]
    [InlineData(WordListLevelKind.LowerRomanDot, "xii.", "xiii.")]
    [InlineData(WordListLevelKind.UpperLetterDot, "L.", "M.")]
    public void WordRichBlocksPreserveResolvedListMarkers(WordListLevelKind kind, string first, string second) {
        using var stream = new MemoryStream();
        using (WordDocument document = WordDocument.Create(stream)) {
            WordList list = document.AddCustomList();
            list.Numbering.AddLevel(new WordListLevel(kind));
            list.Numbering.Levels[0].StartNumberingValue = 12;
            list.AddItem("First item");
            list.AddItem("Second item");
            document.Save();
        }
        stream.Position = 0;

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.Word()
            .ReadDocument(stream, "numbered.docx");

        Assert.Equal(new[] { first, second }, result.Blocks.Where(block => block.Kind == "list-item").Select(block => block.Marker));
    }
}
