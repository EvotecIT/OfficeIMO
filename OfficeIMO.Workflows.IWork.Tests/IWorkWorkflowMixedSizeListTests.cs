using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed partial class IWorkWorkflowTests {
    private static void AssertMixedSizeList(PdfReadDocument pdf, bool edited, bool nativeReference) {
        AssertMixedFontList(pdf, edited, nativeReference ? "Mixed family decimal items" : "Mixed size decimal items");
        var spans = Assert.Single(pdf.Pages).GetTextSpans();
        double[] sizes = [18, 12, 12, 12, 12, 12, 12, 12, 12, 12, 8, 8, 12, 12, 12, 12, 12];
        var markers = spans.Where(span => Regex.IsMatch(span.Text.Trim(), @"^(\d+\.|\(\d+\)|\d+\))$")).ToArray();
        Assert.Equal(edited ? sizes.Concat(new double[] { 8 }) : sizes, markers.Select(span => span.FontSize));
        Assert.Contains(spans, span => span.BaseFont == "Times-Bold" && span.IsBold && span.FontSize == 18);
        Assert.Contains(spans, span => span.BaseFont == "Courier-Oblique" && span.IsItalic && span.FontSize == 24);
        Assert.Contains(spans, span => span.Text == "plain" && span.FontSize == 14 && !span.IsBold && !span.IsItalic);
    }
}
