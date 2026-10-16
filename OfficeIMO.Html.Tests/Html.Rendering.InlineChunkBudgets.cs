using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlInlineChunkBudgetTests {
    [Theory]
    [InlineData("position:relative;z-index:0")]
    [InlineData("background:lime")]
    public void InterruptedInlinePaint_ChargesTheChunkAndOwnerCrossProduct(string ownerStyle) {
        // A small source can create many chunks and independent paint owners.
        // Ordinary source traversal fits; the owner/chunk cross product does not.
        string html = "<style>*{margin:0}body{font:16px/0 Pinned}.owner{" + ownerStyle + "}</style><div><span>"
            + string.Concat(Enumerable.Repeat("<span class='owner'>x</span><span style='display:block;height:0'></span>", 128))
            + "</span></div>";
        var options = new HtmlRenderOptions { MaxLayoutOperations = 20_000 };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf")));

        HtmlDomLimitException error = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options));

        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, error.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), error.LimitSource);
        Assert.Contains("inline paint chunk", error.Message);

        options.MaxLayoutOperations = 200_000;
        HtmlRenderDocument control = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), options);
        Assert.Equal(128, control.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Count());
    }
}
