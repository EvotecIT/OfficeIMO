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
        OfficeIMO.Pdf.PdfConversionReport conversionReport;
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartParagraphSpacing.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartParagraphSpacing.pdf");
        string spacedPdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartParagraphSpacingExplicit.pdf");
        string beforePdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartParagraphSpacingBefore.pdf");
        var options = new WordPdfExportOptions {
            IncludePageNumbers = false,
            PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(420),OfficeIMO.OfficeLength.Points(520)),
            Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(40))
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("BeforeChartSpacingProbe").LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            WordChart chart = document.AddChart("ChartSpacingProbe",false,OfficeIMO.OfficeLength.Pixels(360),OfficeIMO.OfficeLength.Pixels(180));
            document.Paragraphs.Last().LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            chart.AddPie("Passed", 4);
            chart.AddPie("Failed", 2);
            document.AddParagraph("AfterChartSpacingProbe");

            document.Save();
            conversionReport = (document.ConvertToPdfWithEvidence(options).Save(pdfPath)).Report;
        }

        Assert.DoesNotContain(conversionReport.Warnings, warning => warning.Code == "native-body-chart-unsupported");

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        var words = pdf.GetPage(1).GetWords().ToList();
        var chartTitle = Assert.Single(words, word => word.Text == "ChartSpacingProbe");
        var afterChart = Assert.Single(words, word => word.Text == "AfterChartSpacingProbe");

        Assert.True(chartTitle.BoundingBox.Bottom > afterChart.BoundingBox.Top);

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("BeforeChartSpacingProbe").LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            WordChart chart = document.AddChart("ChartSpacingProbe",false,OfficeIMO.OfficeLength.Pixels(360),OfficeIMO.OfficeLength.Pixels(180));
            document.Paragraphs.Last().LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(24D);
            chart.AddPie("Passed", 4);
            chart.AddPie("Failed", 2);
            WordParagraph after = document.AddParagraph("AfterChartSpacingProbe");
            after.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(12D);
            document.Save();
            document.SaveAsPdf(spacedPdfPath, options);
        }

        using PdfPigDocument spacedPdf = PdfPigDocument.Open(spacedPdfPath);
        var spacedText = Assert.Single(spacedPdf.GetPage(1).GetWords(), word => word.Text == "AfterChartSpacingProbe");
        Assert.InRange(afterChart.BoundingBox.Top - spacedText.BoundingBox.Top, 20D, 28D);

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("BeforeChartSpacingProbe").LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(12D);
            WordChart chart = document.AddChart("ChartSpacingProbe",false,OfficeIMO.OfficeLength.Pixels(360),OfficeIMO.OfficeLength.Pixels(180));
            document.Paragraphs.Last().LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(24D);
            document.Paragraphs.Last().LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
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
