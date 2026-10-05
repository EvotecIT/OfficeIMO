using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "body", 6D)]
    [InlineData(true, "body", 6D)]
    [InlineData(false, "columns", 6D)]
    [InlineData(true, "columns", 6D)]
    [InlineData(false, "table", 6D)]
    [InlineData(true, "table", 6D)]
    [InlineData(false, "body", 40D)]
    [InlineData(true, "body", 40D)]
    [InlineData(false, "columns", 40D)]
    [InlineData(true, "columns", 40D)]
    [InlineData(false, "table", 40D)]
    [InlineData(true, "table", 40D)]
    public void SaveAsPdf_ExactLinesKeepTheirBaselinePositionAcrossRunSizes(bool nativeDoc, string frame, double height) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph;
        if (frame == "table") {
            WordTable table = source.AddTable(1, 1);
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.MarginTopWidth = 0; cell.MarginBottomWidth = 0;
            paragraph = cell.Paragraphs[0];
        } else paragraph = source.AddParagraph();
        if (frame == "columns") source.Sections[0].ColumnCount = 2;
        paragraph._paragraph.RemoveAllChildren<Run>();
        paragraph._paragraph.Append(TextLine("A", 8, true), TextLine("B", 48, true), TextLine("C", 8, false));
        paragraph.LineSpacingPoints = height;
        paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
        paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
        using WordDocument document = WordDocument.Load(new MemoryStream(nativeDoc ? source.ToBytes(WordFileFormat.Doc) : source.ToBytes()));
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PageSize = new OfficeIMO.Pdf.PageSize(300, 300), Margins = PageMargins.Uniform(30)
        }));
        var letters = pdf.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "A");
        var second = Assert.Single(letters, letter => letter.Value == "B");
        var third = Assert.Single(letters, letter => letter.Value == "C");
        // Independent Word exports place exact-height baselines at four fifths
        // of the authored line height, including mixed sizes and cell text.
        Assert.Equal(270D - height * .8D, first.StartBaseLine.Y, 3);
        Assert.Equal(height, first.StartBaseLine.Y - second.StartBaseLine.Y, 3);
        Assert.Equal(height, second.StartBaseLine.Y - third.StartBaseLine.Y, 3);
        Assert.Equal(new[] { 8D, 48D, 8D }, new[] { first.PointSize, second.PointSize, third.PointSize });
    }
}
