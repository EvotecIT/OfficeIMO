using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordFileFormat.Docx)]
    [InlineData(WordFileFormat.Doc)]
    public void SaveAsPdf_FitTextCellSpreadsShortTextToTheCellEdge(WordFileFormat format) {
        using var pdf = PdfPigDocument.Open(RenderFitTextCell(format, "FitAB", fitText: true));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        double edge = letters.Single(letter => letter.Value == "Z").EndBaseLine.X;
        var fitted = letters.TakeWhile(letter => letter.Value != "Z").ToArray();

        Assert.Equal("FitAB", string.Concat(fitted.Select(letter => letter.Value)));
        Assert.Equal(edge, fitted[fitted.Length - 1].EndBaseLine.X, 1);
        Assert.All(fitted, letter => Assert.Equal(fitted[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
    }

    [Theory]
    [InlineData(WordFileFormat.Docx)]
    [InlineData(WordFileFormat.Doc)]
    public void SaveAsPdf_FitTextCellCondensesLongTextOntoOneLineWithoutShrinkingTheFont(WordFileFormat format) {
        const string text = "CondensedFitTextCellContentStaysOnOneLine";
        using var control = PdfPigDocument.Open(RenderFitTextCell(format, text, fitText: false));
        using var pdf = PdfPigDocument.Open(RenderFitTextCell(format, text, fitText: true));
        var controlLetters = control.GetPage(1).Letters.Where(letter => letter.Value != "Z").ToArray();
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        double edge = letters.Single(letter => letter.Value == "Z").EndBaseLine.X;
        var fitted = letters.TakeWhile(letter => letter.Value != "Z").ToArray();

        Assert.True(controlLetters.Select(letter => Math.Round(letter.StartBaseLine.Y, 1)).Distinct().Count() > 1);
        Assert.Equal(text, string.Concat(fitted.Select(letter => letter.Value)));
        Assert.All(fitted, letter => Assert.Equal(fitted[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
        Assert.InRange(fitted[fitted.Length - 1].EndBaseLine.X, edge - 1D, edge + 0.05D);
        Assert.All(fitted, letter => Assert.Equal(controlLetters[0].FontSize, letter.FontSize, 3));
    }

    private static byte[] RenderFitTextCell(WordFileFormat format, string text, bool fitText) {
        using WordDocument source = WordDocument.Create();
        WordTable table = source.AddTable(2, 1, WordTableStyle.TableNormal);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        foreach (WordTableRow row in table.Rows) {
            row.Cells[0].Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Points(120D));
        }
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.FitText = fitText;
        cell.Paragraphs[0].Text = text;
        // A right-aligned reference in an identical cell marks the inner right edge.
        WordParagraph edge = table.Rows[1].Cells[0].Paragraphs[0];
        edge.Text = "Z";
        edge.ParagraphAlignment = WordParagraphAlignment.Right;
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(format)));
        return loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
    }
}
