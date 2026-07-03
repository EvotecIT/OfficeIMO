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
        var options = new PdfSaveOptions {
            IncludePageNumbers = false,
            PageSize = new PdfCore.PageSize(420, 520),
            Margins = PdfCore.PageMargins.Uniform(40)
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("ChartSpacingProbe", false, 360, 180);
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
        Assert.True(afterChart.BoundingBox.Bottom < 318D, "A Word chart paragraph should reserve the normal following paragraph clearance.");
    }
}
