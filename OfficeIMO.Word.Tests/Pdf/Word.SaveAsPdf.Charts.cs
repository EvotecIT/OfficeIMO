using System.IO;
using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Reserves_Word_Post_Chart_Paragraph_Spacing() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartParagraphSpacing.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartParagraphSpacing.pdf");
        string spacedPdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartParagraphSpacingExplicit.pdf");
        string beforePdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartParagraphSpacingBefore.pdf");
        var options = new WordToPdfOptions {
            IncludePageNumbers = false,
            PageSize = new PdfCore.PageSize(420, 520),
            Margins = PdfCore.PageMargins.Uniform(40)
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("BeforeChartSpacingProbe").LineSpacingAfterPoints = 0;
            WordChart chart = document.AddChart("ChartSpacingProbe", false, 360, 180);
            document.Paragraphs.Last().LineSpacingAfterPoints = 0;
            chart.AddPie("Passed", 4);
            chart.AddPie("Failed", 2);
            document.AddParagraph("AfterChartSpacingProbe");

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        var words = pdf.GetPage(1).GetWords().ToList();
        var chartTitle = Assert.Single(words, word => word.Text == "ChartSpacingProbe");
        var afterChart = Assert.Single(words, word => word.Text == "AfterChartSpacingProbe");

        Assert.True(chartTitle.BoundingBox.Bottom > afterChart.BoundingBox.Top);

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("BeforeChartSpacingProbe").LineSpacingAfterPoints = 0;
            WordChart chart = document.AddChart("ChartSpacingProbe", false, 360, 180);
            document.Paragraphs.Last().LineSpacingAfterPoints = 24;
            chart.AddPie("Passed", 4);
            chart.AddPie("Failed", 2);
            WordParagraph after = document.AddParagraph("AfterChartSpacingProbe");
            after.LineSpacingBeforePoints = 12;
            document.Save();
            document.SaveAsPdf(spacedPdfPath, options);
        }

        using PdfPigDocument spacedPdf = PdfPigDocument.Open(spacedPdfPath);
        var spacedText = Assert.Single(spacedPdf.GetPage(1).GetWords(), word => word.Text == "AfterChartSpacingProbe");
        Assert.InRange(afterChart.BoundingBox.Top - spacedText.BoundingBox.Top, 20D, 28D);

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("BeforeChartSpacingProbe").LineSpacingAfterPoints = 12;
            WordChart chart = document.AddChart("ChartSpacingProbe", false, 360, 180);
            document.Paragraphs.Last().LineSpacingBeforePoints = 24;
            document.Paragraphs.Last().LineSpacingAfterPoints = 0;
            chart.AddPie("Passed", 4);
            chart.AddPie("Failed", 2);
            document.AddParagraph("AfterChartSpacingProbe");
            document.Save();
            document.SaveAsPdf(beforePdfPath, options);
        }

        using PdfPigDocument beforePdf = PdfPigDocument.Open(beforePdfPath);
        var shiftedChartTitle = Assert.Single(beforePdf.GetPage(1).GetWords(), word => word.Text == "ChartSpacingProbe");
        Assert.InRange(chartTitle.BoundingBox.Top - shiftedChartTitle.BoundingBox.Top, 20D, 28D);

    }
}
