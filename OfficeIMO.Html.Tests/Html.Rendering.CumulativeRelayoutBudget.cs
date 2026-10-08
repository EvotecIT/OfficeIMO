using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void NamedPageRebuildSharesOneLayoutOperationBudget() {
        const string content = "<div style='display:flex'><p>BudgetMarker</p><p>OtherMarker</p></div>";
        var options = new HtmlRenderOptions { Mode=HtmlRenderMode.Paged, PageSize=new OfficePageSize(200d/96d,400d/96d),
            Margins=HtmlRenderMargins.All(0), MaxLayoutOperations=64 };
        string singlePass = "<style>body{margin:0}p{margin:0}</style>" + content;
        Assert.Single(HtmlRenderTestDriver.Render(singlePass, options).Pages);
        // Calibrate the public single-pass admission boundary. Traversal can gain
        // valid checkpoints without turning this cumulative-budget test into an
        // assertion about an incidental number of internal layout operations.
        int lower = 1;
        int upper = options.MaxLayoutOperations;
        while (lower < upper) {
            options.MaxLayoutOperations = lower + (upper - lower) / 2;
            try {
                HtmlRenderTestDriver.Render(singlePass, options);
                upper = options.MaxLayoutOperations;
            } catch (HtmlDomLimitException limit) {
                Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), limit.LimitSource);
                lower = options.MaxLayoutOperations + 1;
            }
        }
        options.MaxLayoutOperations = lower;
        Assert.Single(HtmlRenderTestDriver.Render(singlePass, options).Pages);
        options.PageSize = new OfficePageSize(400d/96d,400d/96d);
        var error = Assert.Throws<HtmlDomLimitException>(() => HtmlRenderTestDriver.Render(
            "<style>@page named{size:200px 400px;margin:0}body{page:named;margin:0}p{margin:0}</style>"+content, options));
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), error.LimitSource);
    }

    [Fact]
    public void NamedPageRebuildSharesBackgroundTileAdmission() {
        string image = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string css = "body{margin:0}.tiles{width:100px;height:100px;background:url(data:image/png;base64,"+image+
            ") repeat;background-size:20px 20px}";
        var options = new HtmlRenderOptions { Mode=HtmlRenderMode.Paged, PageSize=new OfficePageSize(200d/96d,400d/96d),
            Margins=HtmlRenderMargins.All(0), MaxBackgroundImageTiles=25 };
        var single = HtmlRenderTestDriver.Render("<style>"+css+"</style><div class='tiles'></div>", options);
        Assert.DoesNotContain(single.Diagnostics, d => d.Code == HtmlRenderDiagnosticCodes.BackgroundImageTileLimitExceeded);
        options.PageSize = new OfficePageSize(400d/96d,400d/96d);
        var rebuilt = HtmlRenderTestDriver.Render("<style>@page named{size:200px 400px;margin:0}body{page:named}"+
            css+"</style><div class='tiles'></div>", options);
        Assert.Contains(rebuilt.Diagnostics, d => d.Code == HtmlRenderDiagnosticCodes.BackgroundImageTileLimitExceeded);
    }
}
