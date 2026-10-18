using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlFixedHeightOverflow_TableOwnerRetainsCellOverflowWithoutGrowingTheRow(bool anonymous) {
        const string content = "<div style='height:20px'><p>T1</p><p>T2</p><p>T3</p></div>";
        string table = anonymous
            ? "<div style='display:table-row'><div style='display:table-cell'>" + content + "</div></div>"
            : "<table style='border-collapse:collapse;border:0;margin:0'><tr><td style='padding:0;border:0'>"
                + content + "</td></tr></table>";
        string html = FixedHeightOverflowSource("<div style='height:60px'>LEAD</div>" + table + "<p>AFTER</p>");
        HtmlPdfRenderRequestResult output = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf,
                FixedHeightOverflowOptions()));
        HtmlRenderDocument rendered = output.RenderResult.Document;

        Assert.Equal(2, rendered.Pages.Count);
        AssertDescendantSpannerSourceOnce(rendered, "LEAD", "T1", "T2", "T3", "AFTER");
        AssertFixedHeightOverflowPlacement(rendered, "T1", 1, 60D);
        AssertFixedHeightOverflowPlacement(rendered, "T2", 1, 80D);
        AssertFixedHeightOverflowPlacement(rendered, "T3", 2, 0D);
        AssertFixedHeightOverflowPlacement(rendered, "AFTER", 1, 80D);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported
            || diagnostic.Code == HtmlRenderDiagnosticCodes.ConstrainedOverflowPaginationPending);

        PdfCore.PdfReadDocument pdf = PdfCore.PdfReadDocument.Open(output.ToBytes());
        Assert.Equal(2, pdf.Pages.Count);
        // Following flow deliberately overlaps T2. Spatial extraction need not insert a word separator.
        string[] words = pdf.Pages.SelectMany(page => page.GetTextSpans())
            .SelectMany(span => span.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)).ToArray();
        Assert.Equal(new[] { "LEAD", "T1", "T2", "T3", "AFTER" }.OrderBy(word => word, StringComparer.Ordinal),
            words.OrderBy(word => word, StringComparer.Ordinal));
    }

    [Fact]
    public void HtmlFixedHeightOverflow_NestedAllocationOnlySourceDoesNotReportAnOverflowBoundary() {
        HtmlRenderDocument allocationOnly = HtmlRenderTestDriver.Render(FixedHeightOverflowSource(
            "<section style='height:200px'><div style='height:80px'><p>ALLOCATION-ONLY</p></div></section>"),
            FixedHeightOverflowOptions());
        AssertDescendantSpannerSourceOnce(allocationOnly, "ALLOCATION-ONLY");
        Assert.DoesNotContain(allocationOnly.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ConstrainedOverflowPaginationPending);
    }
}
