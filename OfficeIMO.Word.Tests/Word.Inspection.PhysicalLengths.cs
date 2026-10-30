using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordInspectionPhysicalLengthTests {
    [Fact]
    public void InspectionPreservesPhysicalSpacingAndHalfPointFontSizesAcrossSaveAndLoad() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Measured text");
        paragraph.FontSize = OfficeLength.Points(11.5);
        paragraph.IndentationHanging = OfficeLength.Points(3.45);
        paragraph.LineSpacingBefore = OfficeLength.Points(2.35);
        paragraph.LineSpacingAfter = OfficeLength.Zero;
        paragraph.AddTabStop(721);
        document.AddParagraph("Inherited text");
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;

        Check(document.Inspect());
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        using var saved = new MemoryStream(document.ToBytes());
        using WordDocument loaded = WordDocument.Load(saved);
        Check(loaded.Inspect());

        static void Check(WordDocumentInfo snapshot) {
            WordSectionInfo section = Assert.Single(snapshot.Sections);
            Assert.True(section.PageWidth!.Value > OfficeLength.Zero);
            WordParagraphInfo[] paragraphs = section.Elements.OfType<WordParagraphInfo>().ToArray();
            Assert.Equal(OfficeLength.Points(-3.45), paragraphs[0].IndentFirstLine);
            Assert.Equal(OfficeLength.Points(2.35), paragraphs[0].SpaceAbove);
            Assert.Equal(OfficeLength.Zero, paragraphs[0].SpaceBelow);
            Assert.Equal(OfficeLength.Points(36.05), Assert.Single(paragraphs[0].TabStops).Position);
            Assert.Equal(OfficeLength.Points(11.5), Assert.Single(paragraphs[0].Runs).FontSize);
            Assert.Null(Assert.Single(paragraphs[1].Runs).FontSize);
            Assert.Null(paragraphs[1].SpaceAbove);
        }
    }
}
