using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    // Word measures a line by the characters on it: a break ending a line of text adds
    // nothing of its own font, and a line holding only a break takes the break's font.
    [Theory]
    [InlineData("own-run")]
    [InlineData("start-of-run")]
    [InlineData("trailing-own-run")]
    public void SaveAsPdf_LineBreakAfterTextKeepsThatLinesHeight(string placement) {
        double plain = RenderLineBreakBaselines(paragraph => {
            paragraph._paragraph.Append(LineBreakRun("LINEA", null, breakAfter: true), LineBreakRun("LINEB", null));
        })["LINEA"];
        double withLargeBreak = RenderLineBreakBaselines(paragraph => {
            paragraph._paragraph.Append(LineBreakRun("LINEA", null));
            switch (placement) {
                case "own-run": paragraph._paragraph.Append(LineBreakRun(null, 72, breakAfter: true), LineBreakRun("LINEB", null)); break;
                case "start-of-run": paragraph._paragraph.Append(LineBreakRun("LINEB", 72, breakBefore: true)); break;
                default: paragraph._paragraph.Append(LineBreakRun(null, 72, breakAfter: true)); break;
            }
        })["LINEA"];
        Assert.Equal(plain, withLargeBreak, 1);
    }

    [Fact]
    public void SaveAsPdf_LineHoldingOnlyABreakTakesTheBreakFont() {
        Dictionary<string, double> plain = RenderLineBreakBaselines(paragraph => paragraph._paragraph.Append(
            LineBreakRun("LINEA", null, breakAfter: true), LineBreakRun(null, null, breakAfter: true), LineBreakRun("LINEC", null)));
        Dictionary<string, double> large = RenderLineBreakBaselines(paragraph => paragraph._paragraph.Append(
            LineBreakRun("LINEA", null, breakAfter: true), LineBreakRun(null, 72, breakAfter: true), LineBreakRun("LINEC", null)));
        Assert.Equal(plain["LINEA"], large["LINEA"], 1);
        // The empty middle line is a 36pt line, about 26pt taller than an 11pt one.
        Assert.True((large["LINEA"] - large["LINEC"]) - (plain["LINEA"] - plain["LINEC"]) > 20D);
    }

    [Theory]
    [InlineData("own-run")]
    [InlineData("start-of-run")]
    public void SaveAsPdf_LineBreakBesideAFloatingTableKeepsTheLineHeight(string placement) {
        // Lines beside a floating table are framed one at a time; the break must not size those frames either.
        Dictionary<string, double> plain = RenderLineBreakBaselines(paragraph => paragraph._paragraph.Append(
            LineBreakRun("LINEA", null, breakAfter: true), LineBreakRun("LINEB", null, breakAfter: true), LineBreakRun("LINEC", null)), besideFloatingTable: true);
        Dictionary<string, double> large = RenderLineBreakBaselines(paragraph => {
            paragraph._paragraph.Append(LineBreakRun("LINEA", null));
            if (placement == "own-run") paragraph._paragraph.Append(LineBreakRun(null, 72, breakAfter: true), LineBreakRun("LINEB", null, breakAfter: true));
            else paragraph._paragraph.Append(LineBreakRun("LINEB", 72, breakBefore: true), LineBreakRun(null, null, breakAfter: true));
            paragraph._paragraph.Append(LineBreakRun("LINEC", null));
        }, besideFloatingTable: true);
        Assert.Equal(plain["LINEA"], large["LINEA"], 1);
        if (placement == "own-run") Assert.Equal(plain["LINEB"], large["LINEB"], 1);
    }

    private static Run LineBreakRun(string? text, int? halfPoints, bool breakBefore = false, bool breakAfter = false) {
        var run = new Run();
        if (halfPoints.HasValue)
            run.Append(new RunProperties(new FontSize { Val = halfPoints.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) },
                new FontSizeComplexScript { Val = halfPoints.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) }));
        if (breakBefore) run.Append(new Break());
        if (text != null) run.Append(new Text(text) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
        if (breakAfter) run.Append(new Break());
        return run;
    }

    private static Dictionary<string, double> RenderLineBreakBaselines(Action<WordParagraph> fill, bool besideFloatingTable = false) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Intro.");
        if (besideFloatingTable) {
            WordTable table = document.AddTable(1, 1);
            table.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Points(108D));
            table.Rows[0].Height = global::OfficeIMO.OfficeLength.Points(120D);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "FLOATCELL";
            table._tableProperties!.TablePositionProperties = new TablePositionProperties {
                HorizontalAnchor = HorizontalAnchorValues.Margin, VerticalAnchor = VerticalAnchorValues.Text,
                TablePositionXAlignment = HorizontalAlignmentValues.Left, TablePositionY = 1, RightFromText = 180
            };
        }
        fill(document.AddParagraph());
        document.AddParagraph("AFTER text.");
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx)));
        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false })
            .ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        return pdf.GetPage(1).GetWords().Where(word => word.Text.StartsWith("LINE", StringComparison.Ordinal))
            .ToDictionary(word => word.Text, word => word.Letters[0].StartBaseLine.Y);
    }
}
