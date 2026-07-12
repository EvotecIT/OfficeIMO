using OfficeIMO.Html;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_DefaultInputBudgetCanRepresentDefaultInlineResourceBudget() {
        var options = new HtmlRenderOptions();
        long base64Characters = 4L * ((options.MaxTotalResourceBytes + 2L) / 3L);

        Assert.True(options.MaxInputCharacters >= base64Characters);
    }

    [Fact]
    public void HtmlRenderer_RejectsSourceBeforeParsingWhenCharacterBudgetIsExceeded() {
        var options = new HtmlRenderOptions { MaxInputCharacters = 10 };

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderEngine.Render("<p>12345</p>", options));

        Assert.Equal(HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxInputCharacters), exception.LimitSource);
        Assert.Equal(12, exception.Actual);
        Assert.Equal(10, exception.Limit);
        Assert.True(HtmlDiagnosticCatalog.TryGet(exception.Code, out _));
    }

    [Fact]
    public async Task HtmlRenderer_EnforcesDomNodeBudgetForSyncAndAsyncRendering() {
        const string html = "<main><p>one</p><p>two</p></main>";
        var options = new HtmlRenderOptions { MaxHtmlNodes = 4 };

        HtmlDomLimitException syncException = Assert.Throws<HtmlDomLimitException>(() => HtmlRenderEngine.Render(html, options));
        HtmlDomLimitException asyncException = await Assert.ThrowsAsync<HtmlDomLimitException>(() => HtmlRenderEngine.RenderAsync(html, options));

        Assert.Equal(HtmlRenderDiagnosticCodes.NodeLimitExceeded, syncException.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxHtmlNodes), syncException.LimitSource);
        Assert.Equal(5, syncException.Actual);
        Assert.Equal(4, syncException.Limit);
        Assert.Equal(syncException.Code, asyncException.Code);
        Assert.Equal(syncException.Detail, asyncException.Detail);
        Assert.True(HtmlDiagnosticCatalog.TryGet(syncException.Code, out _));
    }

    [Fact]
    public void HtmlRenderOptions_ClonePreservesInputBudgetsAndRejectsInvalidValues() {
        var options = new HtmlRenderOptions { MaxInputCharacters = 1234, MaxHtmlNodes = 4321 };

        HtmlRenderOptions clone = options.Clone();

        Assert.Equal(1234, clone.MaxInputCharacters);
        Assert.Equal(4321, clone.MaxHtmlNodes);
        Assert.Throws<ArgumentOutOfRangeException>(() => HtmlRenderEngine.Render("<p>x</p>", new HtmlRenderOptions { MaxInputCharacters = 0 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => HtmlRenderEngine.Render("<p>x</p>", new HtmlRenderOptions { MaxHtmlNodes = 0 }));
    }
}
