using System.Text;
using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class SvgContentSafetyRasterBudgetTests {
    [Fact]
    public void ExhaustedSamplingBudgetReportsUnavailableVisualInspectionAndRetainsStructuralFindings() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='100' height='100' viewBox='0 0 1000 1000'>" +
            "<rect width='1000' height='1000' fill='white'/>" +
            string.Concat(Enumerable.Repeat("<g transform='scale(1000)' opacity='.5'><rect width='1' height='1' fill='white'/></g>", 64)) +
            "<text font-family='OfficeIMO Shaping Test' font-size='100' fill='white' transform='scale(.1,1)' x='1000' y='100'>A</text>" +
            "<text visibility='hidden' x='10' y='30'>hidden payload</text></svg>");
        var options = new OfficeSvgDrawingReaderOptions { MaximumContentSafetyVisualPixels = 100_000 };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName, ManagedTextShapingTestAssets.CreateFont('A'));

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(svg, readerOptions: options);

        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Contains("bounded SVG drawing surface could not be rendered"));
        Assert.Contains(report.Findings, finding => finding.TextPreview == "hidden payload"
            && finding.Kind == OfficeContentConcealmentKind.HiddenByProperty);
    }
}
