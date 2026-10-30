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
        WordTable table = source.AddTable(1, 1, WordTableStyle.TableGrid);
        // Borderless cells isolate exact-line baselines and clipping from the
        // additional space reserved for table border paint.
        table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Nil, 0U, OfficeIMO.OfficeColor.Black);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(columns ? 2040 : 4800) * 635)));
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(table.Width?.ToNativeValue()) * 635)));
        cell.MarginTop = global::OfficeIMO.OfficeLength.Zero; cell.MarginBottom = global::OfficeIMO.OfficeLength.Zero;
        cell.MarginLeft = global::OfficeIMO.OfficeLength.Zero; cell.MarginRight = global::OfficeIMO.OfficeLength.Zero;
        WordParagraph paragraph = cell.Paragraphs[0];
        paragraph._paragraph.RemoveAllChildren<Run>();
        paragraph._paragraph.Append(TextLine("A", 48, true), TextLine("B", 48, true), TextLine("C", 48, false));
        paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(6D));
        paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero; paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        paragraph.IndentationBefore = global::OfficeIMO.OfficeLength.Points(-30);
        using WordDocument document = WordDocument.Load(new MemoryStream(nativeDoc ? source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }) : source.ToBytes()));
        byte[] bytes = document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(300)), Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30)),
            PdfOptions = new PdfCreateOptions { CompressContentStreams = false }
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
            WordTable table = source.AddTable(1, 1, WordTableStyle.TableGrid);
            table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Nil, 0U, OfficeIMO.OfficeColor.Black);
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.MarginTop = global::OfficeIMO.OfficeLength.Zero; cell.MarginBottom = global::OfficeIMO.OfficeLength.Zero;
            paragraph = cell.Paragraphs[0];
        } else paragraph = source.AddParagraph();
        if (frame == "columns") source.Sections[0].ColumnCount = 2;
        paragraph._paragraph.RemoveAllChildren<Run>();
        paragraph._paragraph.Append(TextLine("A", 8, true), TextLine("B", 48, true), TextLine("C", 8, false));
        paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(height));
        paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero; paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        using WordDocument document = WordDocument.Load(new MemoryStream(nativeDoc ? source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }) : source.ToBytes()));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(300),OfficeIMO.OfficeLength.Points(300)), Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
