using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPrintFittingBudgetTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HtmlPdf_PrintFittingAndLayoutShareTheDecorationWorkBudget(bool asyncPath) {
        var source = HtmlConversionDocument.Parse("<body style='text-decoration:overline 3px;text-decoration-skip-ink:none'><span style='display:none'>"
            + new string('x', 65536) + "</span></body>");
        var options = new HtmlToPdfOptions { MaxLayoutOperations = 300, AutoFitWidePrintContent = false };
        byte[] control = asyncPath ? await source.ToPdfBytesAsync(options) : source.ToPdfBytes(options);
        Assert.Equal("%PDF", System.Text.Encoding.ASCII.GetString(control, 0, 4));

        // Each individual scan fits this limit. Width fitting and layout together
        // must consume one operation ledger rather than restarting the budget.
        options.AutoFitWidePrintContent = true;
        HtmlDomLimitException error = asyncPath
            ? await Assert.ThrowsAsync<HtmlDomLimitException>(() => source.ToPdfBytesAsync(options))
            : Assert.Throws<HtmlDomLimitException>(() => source.ToPdfBytes(options));
        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, error.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), error.LimitSource);
        Assert.Contains("text decoration", error.Message);
    }

    [Theory]
    [InlineData("html")]
    [InlineData("body")]
    public void HtmlPrintFittingDecorationScanHonorsActiveCancellation(string decoratingTag) {
        var source = HtmlConversionDocument.Parse(DecoratedRoot(decoratingTag));
        var document = source.CreateSourceDocumentForConversion();
        HtmlRenderOptions options = PrintFitOptions();
        var styles = HtmlComputedStyleEngine.ComputeForRendering(document, options, HtmlConversionLimits.CreateUntrustedProfile());
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        OperationCanceledException error = Assert.ThrowsAny<OperationCanceledException>(() =>
            HtmlCssPrintFitResolver.TryApplyWideRoot(document, styles, new HtmlCssPageRuleSet(), options,
                cancellation.Token, new HtmlRenderOperationBudget()));
        Assert.Equal(cancellation.Token, error.CancellationToken);
    }

    [Theory]
    [InlineData("html")]
    [InlineData("body")]
    public void HtmlPrintFittingDecorationScanConsumesTheExistingOperationLedger(string decoratingTag) {
        var source = HtmlConversionDocument.Parse(DecoratedRoot(decoratingTag));
        var document = source.CreateSourceDocumentForConversion();
        HtmlRenderOptions options = PrintFitOptions();
        var styles = HtmlComputedStyleEngine.ComputeForRendering(document, options, HtmlConversionLimits.CreateUntrustedProfile());
        var budget = new HtmlRenderOperationBudget();
        budget.ChargeLayoutOperations(99L, options.MaxLayoutOperations, "previous render work");

        HtmlDomLimitException error = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlCssPrintFitResolver.TryApplyWideRoot(document, styles, new HtmlCssPageRuleSet(), options,
                CancellationToken.None, budget));
        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, error.Code);
        Assert.Equal(101L, budget.LayoutOperations);
    }

    private static HtmlRenderOptions PrintFitOptions() => new() {
        Mode = HtmlRenderMode.Paged,
        HonorCssPageRules = true,
        AutoFitWidePrintRoot = true,
        MaxLayoutOperations = 100
    };

    private static string DecoratedRoot(string tag) => tag == "html"
        ? "<html style='text-decoration:overline 3px'><body>Text</body></html>"
        : "<body style='text-decoration:overline 3px'>Text</body>";
}
