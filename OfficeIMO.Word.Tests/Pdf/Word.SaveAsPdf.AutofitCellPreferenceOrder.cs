using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, 0, 72)]
    [InlineData(false, 1, 72)]
    [InlineData(false, 0, 144)]
    [InlineData(false, 1, 144)]
    [InlineData(true, 0, 72)]
    [InlineData(true, 1, 72)]
    [InlineData(true, 0, 144)]
    [InlineData(true, 1, 144)]
    public void SaveAsPdf_AutomaticCellPreferenceRespectsPictureRowOrder(bool reversed, int column, int pictureWidth) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutofitPreferenceOrderControl(document);
        WordTableCell cell = table.Rows[1].Cells[column];
        cell.WidthType = WordTableWidthUnit.Auto;
        cell.Width = 0;
        WordParagraph paragraph = cell.Paragraphs[0];
        paragraph.Text = "A";
        using var image = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 1));
        paragraph.AddImage(image, "marker.png", OfficeIMO.OfficeLength.Pixels(pictureWidth * 96D / 72D), OfficeIMO.OfficeLength.Pixels(24));
        paragraph.AddText("B");
        ReverseAutofitPreferenceRows(table, reversed);
        string source = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;

        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(1, pdf.NumberOfPages);
        var page = pdf.GetPage(1);
        var a = Assert.Single(page.Letters, letter => letter.Value == "A");
        var b = Assert.Single(page.Letters, letter => letter.Value == "B");
        var picture = Assert.Single(page.GetImages());
        Assert.Equal(pictureWidth, picture.BoundingBox.Width, 3);
        Assert.Equal(18D, picture.BoundingBox.Height, 3);
        if (reversed) Assert.True(Math.Abs(a.StartBaseLine.Y - b.StartBaseLine.Y) > 1D);
        else {
            Assert.Equal(a.StartBaseLine.Y, b.StartBaseLine.Y, 3);
            Assert.True(a.BoundingBox.Right <= picture.BoundingBox.Left + .5D);
            Assert.True(picture.BoundingBox.Right <= b.BoundingBox.Left + .5D);
        }
        Assert.Contains("M0", page.Text, StringComparison.Ordinal);
        Assert.Contains("M1", page.Text, StringComparison.Ordinal);
        Assert.Equal(source, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Theory]
    [InlineData(false, WordTableWidthUnit.Auto)]
    [InlineData(true, WordTableWidthUnit.Auto)]
    [InlineData(false, WordTableWidthUnit.Dxa)]
    [InlineData(true, WordTableWidthUnit.Dxa)]
    public void SaveAsPdf_AutomaticCellPreferenceRespectsBreakableTextRowOrder(bool reversed, WordTableWidthUnit widthType) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutofitPreferenceOrderControl(document);
        WordTableCell cell = table.Rows[1].Cells[0];
        cell.WidthType = widthType;
        cell.Width = widthType == WordTableWidthUnit.Auto ? 0 : 1600;
        cell.Paragraphs[0].Text = "Alpha Bravo Charlie Delta";
        ReverseAutofitPreferenceRows(table, reversed);
        string source = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(1, pdf.NumberOfPages);
        const string expected = "AlphaBravoCharlieDelta";
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        int start = Array.FindIndex(letters, letter => letter.Value == "A");
        Assert.True(start >= 0);
        var phrase = letters.Skip(start).Take(expected.Length).ToArray();
        Assert.Equal(expected, string.Concat(phrase.Select(letter => letter.Value)));
        int lines = phrase.Select(letter => Math.Round(letter.StartBaseLine.Y, 3)).Distinct().Count();
        if (!reversed && widthType == WordTableWidthUnit.Auto) Assert.Equal(1, lines);
        else Assert.True(lines > 1);
        Assert.Equal(source, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_AbsoluteMergedCellPreservesAutomaticColumnContentPreference(bool automaticLast) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutofitPreferenceOrderControl(document);
        WordTableCell automatic = table.Rows[1].Cells[0];
        automatic.WidthType = WordTableWidthUnit.Auto;
        automatic.Width = 0;
        automatic.Paragraphs[0].Text = "Alpha Bravo Charlie Delta";
        table.AddRow();
        table.MergeCells(2, 0, 1, 2, copyParagraphs: false);
        WordTableCell merged = table.Rows[2].Cells[0];
        merged.WidthType = WordTableWidthUnit.Dxa;
        merged.Width = table.GridColumnWidth.Sum();
        merged.Paragraphs[0].Text = "MERGED";
        if (automaticLast) {
            W.TableRow[] rows = table._table.Elements<W.TableRow>().ToArray();
            rows[2].Remove();
            table._table.InsertBefore(rows[2], rows[1]);
        }
        string source = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;

        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = Assert.Single(pdf.GetPages());
        var words = page.GetWords().ToArray();
        const string expected = "AlphaBravoCharlieDelta";
        var letters = page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        int start = Array.FindIndex(letters, letter => letter.Value == "A");
        Assert.True(start >= 0);
        var phraseLetters = letters.Skip(start).Take(expected.Length).ToArray();
        Assert.Equal(expected, string.Concat(phraseLetters.Select(letter => letter.Value)));
        Assert.All(phraseLetters, letter => Assert.Equal(phraseLetters[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
        var first = Assert.Single(words, word => word.Text == "M0");
        var second = Assert.Single(words, word => word.Text == "M1");
        // Native Word retains automatic single-column growth in either order:
        // an absolute merged cell constrains the span, not each column's frame.
        Assert.True(second.BoundingBox.Left - first.BoundingBox.Left > table.GridColumnWidth[0] / 20D);
        Assert.Contains("P1", page.Text, StringComparison.Ordinal);
        Assert.Single(words, word => word.Text == "MERGED");
        Assert.Equal(source, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    private static WordTable CreateAutofitPreferenceOrderControl(WordDocument document) {
        WordTable table = CreateAutomaticWidthControl(document, 1600, 1600);
        table.AddRow();
        for (int row = 0; row < 2; row++)
            for (int column = 0; column < 2; column++) {
                WordTableCell cell = table.Rows[row].Cells[column];
                cell.Width = 1600;
                cell.WidthType = WordTableWidthUnit.Dxa;
                cell.MarginLeft = cell.MarginRight = global::OfficeIMO.OfficeLength.Zero;
                cell.MarginTop = cell.MarginBottom = global::OfficeIMO.OfficeLength.Zero;
                WordParagraph paragraph = cell.Paragraphs[0];
                paragraph.Text = (row == 0 ? "M" : "P") + column;
                paragraph.FontFamily = "Arial";
                paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(12);
                paragraph.LineSpacingBefore = paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
                paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Multiple((240) / 240D);
            }
        table.GridColumnWidth = new List<int> { 1600, 1600 };
        return table;
    }

    private static void ReverseAutofitPreferenceRows(WordTable table, bool reverse) {
        if (!reverse) return;
        W.TableRow[] rows = table._table.Elements<W.TableRow>().ToArray();
        rows[1].Remove();
        table._table.InsertBefore(rows[1], rows[0]);
    }
}
