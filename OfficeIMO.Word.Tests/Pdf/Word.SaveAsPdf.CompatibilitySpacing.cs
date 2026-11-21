using W = DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

// Expected spacing follows Microsoft Word 16 PDF exports of the same documents.
public partial class Word {
    [Fact]
    public void SaveAsPdf_AutomaticSpacingInTableCellsSkipsCellEdges() {
        // Inside a cell, automatic spacing is 14 pt between paragraphs and none
        // before the first or after the last paragraph.
        (double first, double second, double after) automatic = MeasureCellSpacing(automatic: true);
        (double first, double second, double after) reference = MeasureCellSpacing(automatic: false);
        Assert.Equal(reference.first, automatic.first, 1);
        Assert.Equal(reference.second, automatic.second, 1);
        Assert.Equal(reference.after, automatic.after, 1);
    }

    [Theory]
    [InlineData("Arial", 12, WordCompatibilityMode.Word2010, true, 14D)]
    [InlineData("Arial", 24, WordCompatibilityMode.Word2010, true, 26D)]
    [InlineData("Courier New", 12, WordCompatibilityMode.Word2010, true, 12D)]
    [InlineData("Arial", 12, WordCompatibilityMode.Word2013, true, 0D)]
    [InlineData("Arial", 12, WordCompatibilityMode.Word2010, false, 0D)]
    public void SaveAsPdf_WordPerfectLineHeightUsesFontSizePlusLeading(string font, int size, WordCompatibilityMode mode,
        bool suppressTopSpacing, double expectedPitch) {
        double pitch = MeasureLinePitch(font, size, mode, suppressTopSpacing);
        // Zero means the option has no effect: the line keeps the font's own line height.
        double expected = expectedPitch > 0D ? expectedPitch : MeasureLinePitch(font, size, WordCompatibilityMode.Word2013, false);
        Assert.Equal(expected, pitch, 1);
    }

    private static (double First, double Second, double After) MeasureCellSpacing(bool automatic) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Intro").LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(0D);
        WordTable table = document.AddTable(1, 1);
        WordParagraph first = table.Rows[0].Cells[0].Paragraphs[0];
        first.Text = "CELLA";
        WordParagraph second = first.AddParagraphAfterSelf();
        second.Text = "CELLB";
        foreach (WordParagraph paragraph in new[] { first, second }) {
            W.ParagraphProperties properties = paragraph._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
            properties.SpacingBetweenLines = automatic
                ? new W.SpacingBetweenLines { Before = "100", BeforeAutoSpacing = true, After = "100", AfterAutoSpacing = true }
                : new W.SpacingBetweenLines { Before = paragraph == first ? "0" : "280", After = paragraph == first ? "0" : "0" };
        }
        WordParagraph after = document.AddParagraph("AFTER");
        after.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(0D);
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx)));
        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var words = pdf.GetPage(1).GetWords().ToList();
        double Baseline(string token) => words.First(word => word.Text == token).Letters[0].StartBaseLine.Y;
        return (Baseline("CELLA"), Baseline("CELLB"), Baseline("AFTER"));
    }

    private static double MeasureLinePitch(string font, int size, WordCompatibilityMode mode, bool suppressTopSpacing) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        if (suppressTopSpacing) {
            W.Settings settings = document._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!;
            W.Compatibility compatibility = settings.GetFirstChild<W.Compatibility>() ?? settings.AppendChild(new W.Compatibility());
            compatibility.InsertAt(new W.SuppressTopSpacingWordPerfect(), 0);
        }
        foreach (string token in new[] { "LINEA", "LINEB" }) {
            WordParagraph paragraph = document.AddParagraph(token);
            paragraph.FontFamily = font;
            paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(size);
            paragraph.LineSpacingBefore = paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(0D);
            paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Multiple(1D);
        }
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx)));
        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var words = pdf.GetPage(1).GetWords().ToList();
        return words.First(word => word.Text == "LINEA").Letters[0].StartBaseLine.Y - words.First(word => word.Text == "LINEB").Letters[0].StartBaseLine.Y;
    }
}
