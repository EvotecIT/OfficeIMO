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
        using var control = PdfPigDocument.Open(RenderFitTextCell(format, "FitAB", fitText: false));
        using var pdf = PdfPigDocument.Open(RenderFitTextCell(format, "FitAB", fitText: true));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        double edge = letters.Single(letter => letter.Value == "Z").EndBaseLine.X;
        var fitted = letters.TakeWhile(letter => letter.Value != "Z").ToArray();

        Assert.Equal("FitAB", string.Concat(fitted.Select(letter => letter.Value)));
        // A fitted line fills its row without wrapping into an extra line.
        Assert.Equal(control.GetPage(1).Letters.Single(letter => letter.Value == "Z").StartBaseLine.Y,
            letters.Single(letter => letter.Value == "Z").StartBaseLine.Y, 2);
        Assert.InRange(fitted[fitted.Length - 1].EndBaseLine.X, edge - 1D, edge + 0.05D);
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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_FitTextCellEndsTheLastVisibleLetterAtTheEdgeBeforeTrailingSpaces(bool separateRun) {
        using var pdf = PdfPigDocument.Open(RenderFitTextCell(WordFileFormat.Docx, separateRun ? "FitAB" : "FitAB  ", fitText: true,
            paragraph => { if (separateRun) paragraph.AddText("  "); }));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        double edge = letters.Single(letter => letter.Value == "Z").EndBaseLine.X;
        var fitted = letters.TakeWhile(letter => letter.Value != "Z").ToArray();

        Assert.Equal("FitAB", string.Concat(fitted.Select(letter => letter.Value)));
        Assert.InRange(fitted[fitted.Length - 1].EndBaseLine.X, edge - 1D, edge + 0.05D);
    }

    [Fact]
    public void SaveAsPdf_FitTextCellWidensASingleLetterToTheCellEdge() {
        using var pdf = PdfPigDocument.Open(RenderFitTextCell(WordFileFormat.Docx, "A", fitText: true));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        double edge = letters.Single(letter => letter.Value == "Z").EndBaseLine.X;

        Assert.InRange(letters.Single(letter => letter.Value == "A").EndBaseLine.X, edge - 1D, edge + 0.05D);
    }

    [Fact]
    public void SaveAsPdf_FitTextCellReplacesAuthoredCharacterSpacing() {
        const string text = "TrackedFitTextCellContent";
        using var pdf = PdfPigDocument.Open(RenderFitTextCell(WordFileFormat.Docx, text, fitText: true,
            paragraph => paragraph.Spacing = global::OfficeIMO.OfficeLength.Points(4D)));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        double edge = letters.Single(letter => letter.Value == "Z").EndBaseLine.X;
        var fitted = letters.TakeWhile(letter => letter.Value != "Z").ToArray();

        Assert.Equal(text, string.Concat(fitted.Select(letter => letter.Value)));
        Assert.All(fitted, letter => Assert.Equal(fitted[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
        Assert.InRange(fitted[fitted.Length - 1].EndBaseLine.X, edge - 1D, edge + 0.05D);
    }

    [Fact]
    public void SaveAsPdf_FitTextCellKeepsTheParagraphMarkLineHeightAfterATrailingBreak() {
        static void TallTrailingBreak(WordParagraph paragraph) {
            paragraph.AddBreak();
            paragraph._paragraph.ParagraphProperties ??= new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties();
            paragraph._paragraph.ParagraphProperties.ParagraphMarkRunProperties =
                new DocumentFormat.OpenXml.Wordprocessing.ParagraphMarkRunProperties(new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "72" });
        }

        using var control = PdfPigDocument.Open(RenderFitTextCell(WordFileFormat.Docx, "Alpha", fitText: false, TallTrailingBreak));
        using var pdf = PdfPigDocument.Open(RenderFitTextCell(WordFileFormat.Docx, "Alpha", fitText: true, TallTrailingBreak));
        double expected = control.GetPage(1).Letters.Single(letter => letter.Value == "Z").StartBaseLine.Y;

        Assert.Equal(expected, pdf.GetPage(1).Letters.Single(letter => letter.Value == "Z").StartBaseLine.Y, 2);
    }

    [Fact]
    public void SaveAsPdf_FitTextCellDoesNotWidenAnAutoFitColumn() {
        byte[] Render(bool fitText) {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(1, 2, WordTableStyle.TableGrid);
            table.LayoutMode = WordTableLayoutMode.AutoFit;
            table.Rows[0].Cells[0].FitText = fitText;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "AB";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Q";
            return document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
                .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        }

        using var control = PdfPigDocument.Open(Render(fitText: false));
        using var pdf = PdfPigDocument.Open(Render(fitText: true));

        Assert.Equal(control.GetPage(1).Letters.Single(letter => letter.Value == "Q").StartBaseLine.X,
            pdf.GetPage(1).Letters.Single(letter => letter.Value == "Q").StartBaseLine.X, 1);
    }

    [Fact]
    public void SaveAsPdf_FitTextCellJoinsParagraphsAcrossAHiddenParagraphMark() {
        using var pdf = PdfPigDocument.Open(RenderFitTextCell(WordFileFormat.Docx, "Alpha", fitText: true, paragraph => {
            paragraph._paragraph.ParagraphProperties ??= new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties();
            paragraph._paragraph.ParagraphProperties.ParagraphMarkRunProperties =
                new DocumentFormat.OpenXml.Wordprocessing.ParagraphMarkRunProperties(new DocumentFormat.OpenXml.Wordprocessing.Vanish());
            paragraph.AddParagraph("Beta");
        }));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        var fitted = letters.TakeWhile(letter => letter.Value != "Z").ToArray();

        Assert.Equal("AlphaBeta", string.Concat(fitted.Select(letter => letter.Value)));
        Assert.All(fitted, letter => Assert.Equal(fitted[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
    }

    [Fact]
    public void SaveAsPdf_FitTextCellFitsLinesBesideALineWithATab() {
        using var pdf = PdfPigDocument.Open(RenderFitTextCell(WordFileFormat.Docx, "Short", fitText: true, paragraph => {
            paragraph.AddBreak();
            paragraph.AddText("Q");
            paragraph.AddTab();
            paragraph.AddText("R");
        }));
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        double edge = letters.Single(letter => letter.Value == "Z").EndBaseLine.X;

        Assert.InRange(letters.Single(letter => letter.Value == "t").EndBaseLine.X, edge - 1D, edge + 0.05D);
        Assert.True(letters.Single(letter => letter.Value == "R").StartBaseLine.X < edge - 20D);
    }

    private static byte[] RenderFitTextCell(WordFileFormat format, string text, bool fitText, Action<WordParagraph>? configure = null) {
        using WordDocument source = WordDocument.Create();
        WordTable table = source.AddTable(2, 1, WordTableStyle.TableNormal);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        foreach (WordTableRow row in table.Rows) {
            row.Cells[0].Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Points(120D));
        }
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.FitText = fitText;
        cell.Paragraphs[0].Text = text;
        configure?.Invoke(cell.Paragraphs[0]);
        // A right-aligned reference in an identical cell marks the inner right edge.
        WordParagraph edge = table.Rows[1].Cells[0].Paragraphs[0];
        edge.Text = "Z";
        edge.ParagraphAlignment = WordParagraphAlignment.Right;
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes(format)));
        return loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
    }
}
