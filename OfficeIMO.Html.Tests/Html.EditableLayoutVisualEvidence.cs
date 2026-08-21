using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Tests;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutVisualEvidenceTests {
    [Fact]
    public void DestinationArtifactsReopenAndRenderNativeEditableGeometry() {
        const string html = "<style>" +
            ".positioned{position:absolute;left:32px;top:200px;width:240px;height:72px;background:#dbeafe;z-index:4}" +
            ".grid{display:grid;grid-template-columns:1fr 1fr;width:300px;height:80px;background:#fef3c7}" +
            "</style><p>Ordinary flow</p><div class='positioned'>Editable positioned</div>" +
            "<div class='grid'>Grid content</div>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);

        HtmlToWordResult wordResult = source.ToWordDocumentResult();
        byte[] docx = SaveWord(wordResult.Value);
        using WordDocument word = WordDocument.Load(
            new MemoryStream(docx),
            new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        byte[] wordPng = word.ToPng(new WordImageExportOptions { BackgroundColor = OfficeColor.White });

        HtmlToRtfResult rtfResult = source.ToRtfDocumentResult();
        byte[] rtf = Encoding.ASCII.GetBytes(rtfResult.Value.ToRtf());
        RtfReadResult reopenedRtf = RtfDocument.Read(Encoding.ASCII.GetString(rtf));

        HtmlToExcelResult excelResult = source.ToExcelDocumentResult(
            new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        byte[] xlsx = SaveExcel(excelResult.Value);
        using ExcelDocument excel = ExcelDocument.Load(
            new MemoryStream(xlsx),
            new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        byte[] excelPng = Assert.Single(excel.Sheets).ToPng(
            new ExcelWorksheetImageExportOptions { Range = "A1:E16" });

        HtmlToPowerPointResult powerPointResult = source.ToPowerPointPresentationResult(
            new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        byte[] pptx = SavePowerPoint(powerPointResult.Value);
        using PowerPointPresentation powerPoint = PowerPointPresentation.Load(
            new MemoryStream(pptx),
            new PowerPointLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        PowerPointSlide powerPointSlide = Assert.Single(powerPoint.Slides);
        PowerPointTextBox ordinaryFlow = Assert.Single(powerPointSlide.TextBoxes, box => box.Text == "Ordinary flow");
        PowerPointTextBox positionedBox = Assert.Single(powerPointSlide.TextBoxes, box => box.Text == "Editable positioned");
        PowerPointTextBox gridBox = Assert.Single(powerPointSlide.TextBoxes, box => box.Text == "Grid content");
        Assert.False(Overlaps(ordinaryFlow, gridBox));
        Assert.False(Overlaps(positionedBox, gridBox));
        byte[] powerPointPng = powerPointSlide.ToPng();

        WriteEvidenceWhenRequested("editable-layout.docx", docx);
        WriteEvidenceWhenRequested("editable-layout.rtf", rtf);
        WriteEvidenceWhenRequested("editable-layout.xlsx", xlsx);
        WriteEvidenceWhenRequested("editable-layout.pptx", pptx);
        WriteEvidenceWhenRequested("editable-layout-word.png", wordPng);
        WriteEvidenceWhenRequested("editable-layout-excel.png", excelPng);
        WriteEvidenceWhenRequested("editable-layout-powerpoint.png", powerPointPng);

        AssertNativeVisual(wordPng, minimumNonWhitePixels: 150, "DOCX");
        AssertNativeVisual(excelPng, minimumNonWhitePixels: 100, "XLSX");
        AssertNativeVisual(powerPointPng, minimumNonWhitePixels: 100, "PPTX");
        Assert.Contains(reopenedRtf.Document.Paragraphs, paragraph =>
            paragraph.Frame.WidthTwips == 3600 &&
            paragraph.ToPlainText().Contains("Editable positioned", StringComparison.Ordinal));
    }

    private static byte[] SaveWord(WordDocument document) {
        using (document) {
            using var stream = new MemoryStream();
            document.Save(stream);
            return stream.ToArray();
        }
    }

    private static byte[] SaveExcel(ExcelDocument document) {
        using (document) {
            using var stream = new MemoryStream();
            document.Save(stream);
            return stream.ToArray();
        }
    }

    private static byte[] SavePowerPoint(PowerPointPresentation presentation) {
        using (presentation) {
            using var stream = new MemoryStream();
            presentation.Save(stream);
            return stream.ToArray();
        }
    }

    private static void AssertNativeVisual(byte[] png, int minimumNonWhitePixels, string destination) {
        OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(
            png,
            destination + " editable-layout evidence is not a valid PNG.");
        Assert.True(image.Width > 0 && image.Height > 0);
        Assert.True(
            VisualBaselineTestSupport.CountNonWhiteVisiblePixels(image) >= minimumNonWhitePixels,
            destination + " editable-layout evidence did not contain enough rendered content.");
        Assert.True(
            CountPixelsNear(image, OfficeColor.FromRgb(219, 234, 254), tolerance: 12) >= 40,
            destination + " editable-layout evidence did not render the native positioned-region fill.");
    }

    private static int CountPixelsNear(OfficeRasterImage image, OfficeColor expected, int tolerance) {
        int count = 0;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor actual = image.GetPixel(x, y);
                if (Math.Abs(actual.R - expected.R) <= tolerance &&
                    Math.Abs(actual.G - expected.G) <= tolerance &&
                    Math.Abs(actual.B - expected.B) <= tolerance) {
                    count++;
                }
            }
        }
        return count;
    }

    private static bool Overlaps(PowerPointTextBox first, PowerPointTextBox second) =>
        first.LeftPoints < second.LeftPoints + second.WidthPoints
        && first.LeftPoints + first.WidthPoints > second.LeftPoints
        && first.TopPoints < second.TopPoints + second.HeightPoints
        && first.TopPoints + first.HeightPoints > second.TopPoints;

    private static void WriteEvidenceWhenRequested(string fileName, byte[] bytes) {
        string? folder = Environment.GetEnvironmentVariable("OFFICEIMO_HTML_LAYOUT_EVIDENCE_DIR");
        if (string.IsNullOrWhiteSpace(folder)) return;
        Directory.CreateDirectory(folder);
        File.WriteAllBytes(Path.Combine(folder, fileName), bytes);
    }
}
