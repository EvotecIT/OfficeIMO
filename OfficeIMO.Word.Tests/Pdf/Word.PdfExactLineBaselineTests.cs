using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void SaveAsPdf_TableTextKeepsTheCellClipWithNegativeIndents(bool nativeDoc, bool columns) {
        using WordDocument source = WordDocument.Create();
        if (columns) source.Sections[0].ColumnCount = 2;
        WordTable table = source.AddTable(1, 1);
        // Borderless cells isolate exact-line baselines and clipping from the
        // additional space reserved for table border paint.
        table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Nil, 0U, OfficeIMO.Drawing.OfficeColor.Black);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = columns ? 2040 : 4800;
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.WidthType = WordTableWidthUnit.Dxa; cell.Width = table.Width;
        cell.MarginTopWidth = 0; cell.MarginBottomWidth = 0;
        cell.MarginLeftWidth = 0; cell.MarginRightWidth = 0;
        WordParagraph paragraph = cell.Paragraphs[0];
        paragraph._paragraph.RemoveAllChildren<Run>();
        paragraph._paragraph.Append(TextLine("A", 48, true), TextLine("B", 48, true), TextLine("C", 48, false));
        paragraph.LineSpacingPoints = 6D;
        paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
        paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
        paragraph.IndentationBeforePoints = -30;
        using WordDocument document = WordDocument.Load(new MemoryStream(nativeDoc ? source.ToBytes(WordFileFormat.Doc) : source.ToBytes()));
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PageSize = new OfficeIMO.Pdf.PageSize(300, 300), Margins = PageMargins.Uniform(30),
            PdfOptions = new PdfOptions { CompressContentStreams = false }
        });
        // Word retains the authored negative glyph positions in the PDF but
        // clips their painting to the cell, even with exact-height overlap.
        string content = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains(columns ? "30 252 102 18 re W n" : "30 252 240 18 re W n", content);
        using var pdf = PdfPigDocument.Open(bytes);
        var first = Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "A");
        Assert.Equal(0D, first.StartBaseLine.X, 3);
        Assert.Equal(265.2D, first.StartBaseLine.Y, 3);
    }

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
            table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Nil, 0U, OfficeIMO.Drawing.OfficeColor.Black);
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
