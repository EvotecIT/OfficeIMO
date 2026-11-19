using System.Text;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_Paged_ChargesForcedBreakCursorsByWorkNotByDocumentLength() {
        // A thousand sections inside one container share one forced-break list, so every printed
        // page consults the same list. The cursor cost per page must not grow with the document.
        var html = new StringBuilder("<style>section{break-after:page}section:last-child{break-after:auto}</style><main>");
        for (int index = 1; index <= 1000; index++) html.Append("<section>Article ").Append(index).Append("</section>");
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html.Append("</main>").ToString());
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            MaxPageCount = 1000,
            // One hundred operations per page: a cursor charged by list length needs over a million.
            MaxLayoutOperations = 100_000
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);

        Assert.Equal(1000, rendered.Pages.Count);
        Assert.Contains("Article 1000", rendered.Text, StringComparison.Ordinal);

        // The limit still stops the same document when the budget cannot cover ten operations per page.
        options.MaxLayoutOperations = 10_000;
        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render(document, options));
        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), exception.LimitSource);
    }
}
