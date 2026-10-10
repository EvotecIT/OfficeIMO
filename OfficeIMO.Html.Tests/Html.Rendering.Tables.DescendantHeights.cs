using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlTableCellHeightDescendantTests {
    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, "block", 50D, 150D)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, "block", 80D, 120D)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged, "inline-block", 80D, 120D)]
    public void FinalCellHeightSizesOrdinaryDescendantsInTheEffectiveMedia(HtmlRenderIntentProfile profile, string display, double first, double second) {
        string html = "<style>@media print{#first{height:25%}#second{height:75%}}"
            + "@media screen{#first{height:40%}#second{height:60%}}</style>"
            + "<table style='height:200px'><tr><td id='first'><div id='fill' style='display:" + display + ";height:100%;background:lime'>Paid</div></td></tr>"
            + "<tr><td id='second'><div id='detail' style='height:100%;background:blue'>Due</div></td></tr></table>" + Following;
        HtmlRenderDocument rendered = Render(html, profile);

        Assert.Equal(first, Shape(rendered, "div#fill").Height, 3);
        Assert.Equal(second, Shape(rendered, "div#detail").Height, 3);
        Assert.Equal(200D, Shape(rendered, "div#following").Y, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("height:100%", 100D)]
    [InlineData("height:calc(var(--fill) - 10px)", 90D)]
    [InlineData("height:20px;min-height:100%", 100D)]
    [InlineData("height:80px;max-height:50%", 50D)]
    [InlineData("block-size:var(--fill);height:30px;block-size:calc(100% - 10px)", 90D)]
    public void EffectiveHeightConstraintsUseTheUsedCellHeight(string sizing, double expected) {
        string html = "<table style='height:200px;--fill:100%'><tr><td style='height:50%'>"
            + "<div id='fill' style='background:lime;" + sizing + "'>Paid</div></td></tr>"
            + "<tr><td style='height:50%'>Due</td></tr></table>" + Following;
        HtmlRenderDocument rendered = Render(html);

        Assert.Equal(expected, Shape(rendered, "div#fill").Height, 3);
        Assert.Equal(200D, Shape(rendered, "div#following").Y, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("content-box", "120px", "block", "height:240px;max-block-size:calc(var(--cap) + 0px)", 240D, 242D)]
    [InlineData("border-box", "122px", "block", "height:240px;max-block-size:calc(var(--cap) + 0px)", 240D, 242D)]
    [InlineData("content-box", "120px", "inline-block", "height:1px;max-height:30px;min-block-size:var(--cap)", 120D, 122D)]
    [InlineData("border-box", "122px", "inline-block", "height:1px;max-height:30px;min-block-size:var(--cap)", 120D, 122D)]
    public void AbsoluteCellConstraintsMeasureIntrinsicContentBeforeFinalPercentageSizing(string sizing, string height, string display, string constraints, double expected, double following) {
        HtmlRenderDocument rendered = Render("<table style='--cap:100%'><tr><td style='height:" + height + ";padding:1px;border:0;box-sizing:" + sizing + "'>"
            + "<div id='fill' style='display:" + display + ";background:lime;" + constraints + "'></div></td></tr></table>" + Following);

        Assert.Equal(expected, Shape(rendered, "div#fill").Height, 3);
        Assert.Equal(following, Shape(rendered, "div#following").Y, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("height:200%")]
    [InlineData("height:1px;min-height:200%")]
    public void PercentageOverflowDoesNotFeedBackIntoTheAbsoluteCellMinimum(string constraints) {
        HtmlRenderDocument rendered = Render("<table><tr><td style='height:120px'><span style='height:20px'>"
            + "<div id='fill' style='display:inline-block;width:40px;background:lime;" + constraints + "'></div></span></td></tr></table>" + Following);

        Assert.Equal(240D, Shape(rendered, "div#fill").Height, 3);
        Assert.Equal(120D, Shape(rendered, "div#following").Y, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("content-box")]
    [InlineData("border-box")]
    public void FinalBasisDeductsCellPaddingAndBordersOnce(string sizing) {
        HtmlRenderDocument rendered = Render("<table style='height:200px'><tr><td style='height:100%;padding:10px;border:2px solid red;box-sizing:" + sizing + "'>"
            + "<div id='fill' style='height:100%;background:lime'>Paid</div></td></tr></table>" + Following);

        Assert.Equal(12D, Shape(rendered, "div#fill").Y, 3);
        Assert.Equal(176D, Shape(rendered, "div#fill").Height, 3);
        Assert.Equal(200D, Shape(rendered, "div#following").Y, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void AutoCellsAndNestedPercentageChainsUseTheFinalDefiniteTableAllocation() {
        HtmlRenderDocument rendered = Render("<table style='height:200px'><tr><td><div style='height:100%'>"
            + "<div id='fill' style='height:50%;background:lime'>Paid</div></div></td></tr><tr><td>Due</td></tr></table>" + Following);

        Assert.Equal(50D, Shape(rendered, "div#fill").Height, 3);
        Assert.Equal(200D, Shape(rendered, "div#following").Y, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void CssTableCellsUseTheSameFinalAllocation() {
        HtmlRenderDocument rendered = Render("<div style='display:table;height:200px;width:240px'><div style='display:table-row'>"
            + "<div style='display:table-cell;height:25%'><div id='fill' style='height:100%;background:lime'>Paid</div></div></div>"
            + "<div style='display:table-row'><div style='display:table-cell;height:75%'>Due</div></div></div>" + Following);

        Assert.Equal(50D, Shape(rendered, "div#fill").Height, 3);
        Assert.Equal(200D, Shape(rendered, "div#following").Y, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("inline")]
    [InlineData("contents")]
    public void NonFormattingWrappersDoNotBreakTheCellHeightBasis(string display) {
        HtmlRenderDocument rendered = Render("<table style='height:200px'><tr><td><span style='display:" + display + "'>"
            + "<span id='fill' style='display:inline-block;height:100%;background:lime'>Paid</span></span></td></tr></table>" + Following);

        Assert.Equal(200D, Shape(rendered, "span#fill").Height, 3);
        Assert.Equal(200D, Shape(rendered, "div#following").Y, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("inline", "height:20px", HtmlRenderIntentProfile.PrintPaged)]
    [InlineData("contents", "height:20px", HtmlRenderIntentProfile.PrintPaged)]
    [InlineData("inline", "height:auto", HtmlRenderIntentProfile.ScreenMediaPaged)]
    [InlineData("contents", "height:30px;block-size:calc(var(--ignored) + 0px)", HtmlRenderIntentProfile.ScreenSnapshotPaged)]
    public void WrappedReplacedPercentagesUseTheAbsoluteCellBasisDuringFirstLayout(string display, string wrapperSizing, HtmlRenderIntentProfile profile) {
        HtmlRenderDocument rendered = Render("<table style='--ignored:20px;--fill:100%'><tr><td style='height:120px'>"
            + "<span style='display:" + display + ";" + wrapperSizing + "'>"
            + "<img id='image' style='display:block;height:10px;block-size:var(--fill);width:20px' src='data:image/png;base64," + Pixel + "'>"
            + "</span></td></tr></table>" + Following, profile);

        Assert.Equal(120D, Assert.Single(Visuals(rendered).OfType<HtmlRenderImage>()).Height, 3);
        Assert.Equal(120D, Shape(rendered, "div#following").Y, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("inline", "height:120px")]
    [InlineData("contents", "height:var(--ignored)")]
    [InlineData("inline", "height:20px;block-size:calc(var(--ignored) + 0px);min-height:240px")]
    [InlineData("contents", "block-size:120px;height:auto")]
    public void IgnoredWrapperHeightsDoNotMakeAnAutoHeightCellDefinite(string display, string wrapperSizing) {
        HtmlRenderDocument rendered = Render("<table style='--ignored:120px;--fill:100%'><tr><td>"
            + "<span style='display:" + display + ";" + wrapperSizing + "'>"
            + "<span id='fill' style='display:inline-block;block-size:calc(var(--fill) + 0px);background:lime'>Paid</span>"
            + "</span></td></tr></table>" + Following);

        Assert.Equal(20D, Shape(rendered, "span#fill").Height, 3);
        Assert.Equal(20D, Shape(rendered, "div#following").Y, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void AnActualInlineBlockKeepsItsIndependentDefiniteHeightBasis() {
        HtmlRenderDocument rendered = Render("<table><tr><td style='height:120px'><span style='display:inline-block;height:20px'>"
            + "<img style='display:block;height:100%;width:20px' src='data:image/png;base64," + Pixel + "'></span></td></tr></table>" + Following);

        Assert.Equal(20D, Assert.Single(Visuals(rendered).OfType<HtmlRenderImage>()).Height, 3);
        Assert.Equal(120D, Shape(rendered, "div#following").Y, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("position:absolute;display:block")]
    [InlineData("display:flex")]
    public void SpecializedPercentageContentCannotUseCellHeightEqualityAsItsQualification(string formatting) {
        HtmlRenderDocument rendered = Render("<table><tr><td style='height:120px'><span style='height:20px'>"
            + "<div id='fill' style='" + formatting + ";height:100%;background:lime'>Paid</div></span></td></tr></table>" + Following);
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);

        Assert.Equal("div#fill", diagnostic.Source);
        Assert.Contains("height=100%", diagnostic.Detail);
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void PositionedGeneratedPercentagesInAnAbsoluteCellRetainTheirStrictBoundary() {
        HtmlRenderDocument rendered = Render("<style>#cell::before{content:'Paid';position:absolute;display:block;height:100%;background:lime}</style>"
            + "<table><tr><td id='cell' style='height:120px'></td></tr></table>" + Following);
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);

        Assert.Equal("td#cell::before", diagnostic.Source);
        Assert.Contains("height=100%", diagnostic.Detail);
        Assert.Contains("generated", diagnostic.Detail);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void OrdinaryGeneratedCellContentUsesTheFinalHeightAndRetainsTextOrder() {
        HtmlRenderDocument rendered = Render("<style>#cell::before{content:'Paid';display:block;height:100%;background:lime}</style>"
            + "<table style='height:200px'><tr><td id='cell'></td></tr></table>" + Following);

        Assert.Equal(200D, Shape(rendered, "td#cell::before").Height, 3);
        Assert.Equal(new[] { "Paid", "Following" }, Text(rendered).OrderBy(x => x.LogicalTextOrder).Select(x => x.Text));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void PositionedGeneratedPercentageContentReportsItsOwnSourceProperty() {
        HtmlRenderDocument rendered = Render("<style>#cell::before{content:'Paid';position:absolute;display:block;height:100%;background:lime}</style>"
            + "<table style='height:200px'><tr><td id='cell'></td></tr></table>" + Following);
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);

        Assert.Equal("td#cell::before", diagnostic.Source);
        Assert.Contains("height=100%", diagnostic.Detail);
        Assert.Contains("generated", diagnostic.Detail);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Theory]
    [InlineData("", "", 20D)]
    [InlineData("", "height:120px", 120D)]
    [InlineData("height:200px", "height:40px", 200D)]
    public void AbsoluteMinimaAndLegitimateIndefinitePercentagesRemainIntact(string tableHeight, string cellHeight, double expected) {
        HtmlRenderDocument rendered = Render("<table style='" + tableHeight + "'><tr><td style='" + cellHeight + "'>"
            + "<div id='fill' style='height:100%;background:lime'>Paid</div></td></tr></table>" + Following);

        Assert.Equal(expected, Shape(rendered, "div#fill").Height, 3);
        Assert.Equal(expected, Shape(rendered, "div#following").Y, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void AnAlreadyDefiniteAbsoluteCellPreservesItsReplacedPercentageContent() {
        const string pixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNg+P//HwAF/gL9HjcXBgAAAABJRU5ErkJggg==";
        HtmlRenderDocument rendered = Render("<table><tr><td style='height:120px'><img id='image' style='display:block;height:100%;width:20px' src='data:image/png;base64," + pixel + "'></td></tr></table>" + Following);

        Assert.Equal(120D, Assert.Single(Visuals(rendered).OfType<HtmlRenderImage>()).Height, 3);
        Assert.Equal(120D, Shape(rendered, "div#following").Y, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public void ContentMinimumConstrainsRowsBeforePercentageContentIsResolved() {
        HtmlRenderDocument rendered = Render("<table style='height:200px'><tr><td style='height:25%'><div id='fill' style='height:100%;background:lime'>"
            + "A<br>B<br>C<br>D</div></td></tr><tr><td style='height:75%'><div id='detail' style='height:100%;background:blue'>Due</div></td></tr></table>" + Following);

        Assert.Equal(80D, Shape(rendered, "div#fill").Height, 3);
        Assert.Equal(120D, Shape(rendered, "div#detail").Height, 3);
        Assert.Equal(200D, Shape(rendered, "div#following").Y, 3);
        Assert.Equal(new[] { "A", "B", "C", "D", "Due", "Following" }, Text(rendered).Select(x => x.Text));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("display:flex;")]
    [InlineData("writing-mode:vertical-rl;")]
    public void UnqualifiedFormattingReportsTheSourceAndEffectivePercentageProperty(string formatting) {
        HtmlRenderDocument rendered = Render("<table style='height:200px'><tr><td><div id='fill' style='" + formatting
            + "height:100%;background:lime'>Paid</div></td></tr></table>" + Following);
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);

        Assert.Equal("div#fill", diagnostic.Source);
        Assert.Contains("height=100%", diagnostic.Detail);
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void RowspanPercentageContentKeepsAnExplicitStrictLossBoundary() {
        HtmlRenderDocument rendered = Render("<table style='height:200px'><tr><td rowspan='2'><div id='fill' style='height:100%;background:lime'>Paid</div></td><td>A</td></tr>"
            + "<tr><td>B</td></tr></table>" + Following);
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);

        Assert.Equal("div#fill", diagnostic.Source);
        Assert.Contains("rowspan", diagnostic.Detail);
        Assert.Equal(200D, Shape(rendered, "div#following").Y, 3);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void AnonymousCellPercentageContentReportsItsSourceProperty() {
        HtmlRenderDocument rendered = Render("<div style='display:table;height:200px;width:240px'><div style='display:table-row'>"
            + "<div id='fill' style='height:100%;background:lime'>Paid</div></div></div>" + Following);
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);

        Assert.Equal("div#fill", diagnostic.Source);
        Assert.Contains("height=100%", diagnostic.Detail);
        Assert.Contains("anonymous", diagnostic.Detail);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Theory]
    [InlineData("height:100%;height:20px")]
    [InlineData("height:calc(100% + invalid)")]
    public void AnIneffectivePercentageDoesNotCreateAnUnsupportedHeightLoss(string sizing) {
        HtmlRenderDocument rendered = Render("<table style='height:200px'><tr><td><div id='fill' style='display:flex;" + sizing + ";background:lime'>Paid</div></td></tr></table>" + Following);

        Assert.Equal(20D, Shape(rendered, "div#fill").Height, 3);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("height:var(--share)", "25%")]
    [InlineData("height:calc(25% + 1px - 1px)", "calc(25% + 1px - 1px)")]
    public void EffectiveRowGroupPercentagesReportTheirUnimplementedAllocation(string sizing, string effective) {
        HtmlRenderDocument rendered = Render("<table style='height:200px;--share:25%'><tbody id='group' style='" + sizing
            + "'><tr><td>Paid</td></tr><tr><td>Due</td></tr></tbody><tbody><tr><td>Other</td></tr></tbody></table>" + Following);
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);

        Assert.Equal("tbody#group", diagnostic.Source);
        Assert.Contains("height=" + effective, diagnostic.Detail);
        Assert.Contains("row-group", diagnostic.Detail);
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Theory]
    [InlineData("height:auto")]
    [InlineData("height:25%;height:initial")]
    [InlineData("height:25%;height:unset")]
    [InlineData("height:calc(25% + 10px)")]
    [InlineData("height:calc(25% + invalid)")]
    [InlineData("height:0%")]
    public void IndefiniteOrIneffectiveRowGroupPercentagesDoNotReportHeightLoss(string sizing) {
        HtmlRenderDocument rendered = Render("<table style='height:200px'><tbody style='" + sizing + "'><tr><td>Paid</td></tr></tbody></table>" + Following);

        rendered.RequireNoLoss();
    }

    [Fact]
    public void ARowGroupPercentageInAnAutoHeightTableKeepsItsIndefiniteBasis() {
        HtmlRenderDocument rendered = Render("<table><tbody style='height:25%'><tr><td>Paid</td></tr></tbody></table>" + Following);

        Assert.Equal(20D, Shape(rendered, "div#following").Y, 3);
        rendered.RequireNoLoss();
    }

    [Fact]
    public async Task RelayoutReusesResolvedResourcesAndFontDiagnostics() {
        const string pixel = "iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAAE0lEQVR4nGP4z8AARGAChv6DEQBwqQn3AyNm5wAAAABJRU5ErkJggg==";
        var requests = new List<string>();
        var options = Options();
        options.ResourceResolver = (request, cancellationToken) => {
            cancellationToken.ThrowIfCancellationRequested();
            requests.Add(request.Uri.AbsoluteUri);
            HtmlResolvedResource? resource = request.Kind switch {
                HtmlResourceKind.Stylesheet => new HtmlResolvedResource(System.Text.Encoding.UTF8.GetBytes(
                    "@font-face{font-family:Missing;src:url(missing.ttf)}#fill{height:100%;font-family:Missing,Arial;background:lime url(tile.png) right bottom no-repeat}"), "text/css"),
                HtmlResourceKind.Image => new HtmlResolvedResource(Convert.FromBase64String(pixel), "image/png"),
                _ => null
            };
            return Task.FromResult(resource);
        };
        var source = HtmlConversionDocument.Parse(Source("<link rel='stylesheet' href='https://assets.example.test/site.css'>"
            + "<table style='height:200px'><tr><td style='height:50%'><span style='height:20px'><div id='fill'>Paid</div></span></td></tr><tr><td style='height:50%'>Due</td></tr></table>" + Following));
        HtmlRenderDocument rendered = await HtmlRenderEngine.RenderAsync(source, options);

        Assert.Equal(100D, Shape(rendered, "div#fill").Height, 3);
        Assert.Equal(3, requests.Count);
        Assert.Equal(3, requests.Distinct().Count());
        Assert.Single(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.FontFaceUnavailable);
        HtmlRenderImage background = Assert.Single(Visuals(rendered).OfType<HtmlRenderImage>());
        Assert.Equal(98D, background.Y, 3);
        Assert.DoesNotContain(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);
    }

    [Fact]
    public void CellRelayoutSharesTheRenderOperationBudget() {
        var options = Options();
        HtmlRenderDocument Execute(string height) => HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(Source(
            "<table style='height:200px'><tr><td><div id='fill' style='height:" + height + ";background:lime'>Paid</div></td></tr></table>" + Following)),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, options: options)).Document;
        int lower = 1, upper = 512;
        while (lower < upper) {
            options.MaxLayoutOperations = lower + (upper - lower) / 2;
            try { Execute("20px"); upper = options.MaxLayoutOperations; }
            catch (HtmlDomLimitException error) {
                Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), error.LimitSource);
                lower = options.MaxLayoutOperations + 1;
            }
        }
        options.MaxLayoutOperations = lower;
        Execute("20px").RequireNoLoss();
        HtmlDomLimitException limit = Assert.Throws<HtmlDomLimitException>(() => Execute("100%"));
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), limit.LimitSource);
    }

    [Fact]
    public void RelayoutPreservesLogicalTextOrderCountersAndLinksBesideUnchangedCells() {
        string html = "<style>table{counter-reset:item}.count::before{counter-increment:item;content:counter(item) ': '}</style>"
            + "<table style='height:200px'><tr><td style='height:50%'><span style='height:20px'><div id='fill' class='count' style='height:100%;background:lime'><a href='https://example.com/paid'>Paid</a></div></span></td><td>First note</td></tr>"
            + "<tr><td style='height:50%'><div class='count'>Due</div></td><td>Second note</td></tr></table>" + Following;
        var result = HtmlConversionDocument.Parse(Source(html)).RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, Options()));
        HtmlRenderDocument rendered = result.RenderResult.Document;
        string text = string.Concat(Text(rendered).OrderBy(x => x.LogicalTextOrder).Select(x => x.Text));

        Assert.Equal("1: PaidFirst note2: DueSecond noteFollowing", text);
        var pdf = OfficeIMO.Pdf.PdfReadDocument.Open(result.ToBytes());
        Assert.Single(pdf.Pages[0].GetLinkAnnotations());
        Assert.Contains("Paid", pdf.ExtractText());
        result.Output.RequireNoLoss();
    }

    [Fact]
    public void PagedReportRetainsFinalFillHeightsRepeatedGroupsAndFollowingContent() {
        string rows = string.Concat(Enumerable.Range(1, 4).Select(i => "<tr style='height:20%'><td><div id='fill" + i + "' style='height:100%;background:lime'>Item" + i + "</div></td></tr>"));
        string html = "<table style='height:400px'><thead><tr style='height:5%'><td>Header</td></tr></thead><tbody>" + rows
            + "</tbody><tfoot><tr style='height:15%'><td>Footer</td></tr></tfoot></table>" + Following;
        var options = Options();
        options.PageSize = new OfficePageSize(600D / 96D, 260D / 96D);
        var result = HtmlConversionDocument.Parse(Source(html)).RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));
        HtmlRenderDocument rendered = result.RenderResult.Document;

        Assert.Equal(2, rendered.Pages.Count);
        foreach (int i in Enumerable.Range(1, 4)) {
            var fragments = Visuals(rendered).OfType<HtmlRenderShape>().Where(x => x.Source == "div#fill" + i).ToArray();
            Assert.NotEmpty(fragments);
            Assert.All(fragments, shape => Assert.Equal(80D, shape.Height, 3));
        }
        Assert.Equal(220D, Shape(rendered, "div#following").Y, 3);
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(result.ToBytes()).ExtractText();
        foreach (string token in new[] { "Header", "Footer", "Item1", "Item2", "Item3", "Item4", "Following" }) Assert.Contains(token, text);
        result.Output.RequireNoLoss();
    }

    private const string Pixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNg+P//HwAF/gL9HjcXBgAAAABJRU5ErkJggg==";
    private const string Following = "<div id='following' style='height:20px;background:red'>Following</div>";
    private static string Source(string html) => "<!doctype html><style>html,body{margin:0;padding:0;font:16px/20px Arial}table{margin:0;width:240px;border-spacing:0;table-layout:fixed}td{padding:0;vertical-align:top}</style>" + html;
    private static HtmlRenderOptions Options() => new() { ViewportWidth = 600, ViewportHeight = 300, PageSize = new OfficePageSize(600D / 96D, 300D / 96D), Margins = HtmlRenderMargins.All(0), UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser, HonorCssPageRules = false };
    private static HtmlRenderDocument Render(string html, HtmlRenderIntentProfile profile = HtmlRenderIntentProfile.PrintPaged) => HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(Source(html)), HtmlRenderRequest.Create(profile, options: Options())).Document;
    private static IEnumerable<HtmlRenderVisual> Visuals(HtmlRenderDocument rendered) => rendered.Pages.SelectMany(page => Enumerate(page.Scene));
    private static IEnumerable<HtmlRenderVisual> Enumerate(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IReadOnlyList<HtmlRenderVisual>? children = visual switch {
                HtmlRenderSemanticGroup group => group.Visuals,
                HtmlRenderClipGroup group => group.Visuals,
                HtmlRenderPathClipGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals,
                HtmlRenderLogicalTextGroup group => group.Visuals,
                _ => null
            };
            if (children != null) foreach (HtmlRenderVisual child in Enumerate(children)) yield return child;
        }
    }
    private static HtmlRenderShape Shape(HtmlRenderDocument rendered, string source) => Assert.Single(Visuals(rendered).OfType<HtmlRenderShape>(), x => x.Source == source && x.Shape.FillColor.HasValue);
    private static IEnumerable<HtmlRenderText> Text(HtmlRenderDocument rendered) => Visuals(rendered).OfType<HtmlRenderText>();
}
