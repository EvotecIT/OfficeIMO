using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, false, 5_000)]
    [InlineData(true, false, 5_000)]
    [InlineData(false, true, 25_000)]
    [InlineData(true, true, 25_000)]
    public void HtmlPageMarginSizing_RepeatedNestedRunningMeasurementConsumesTheOperationBudget(
        bool table, bool configuredProvider, int budget) {
        string heading = string.Join(" ", Enumerable.Repeat(new string('H', 16), 128));
        string running = table
            ? "<table id='running'><tbody><tr><td>" + heading + "</td></tr></tbody></table>"
            : "<div id='running'><p>" + heading + "</p></div>";
        string body = string.Join("", Enumerable.Range(1, 20).Select(page =>
            "<p" + (page == 1 ? "" : " style='break-before:page'") + ">B</p>"));
        string Source(bool margin) => "<style>html,body,p,div,table,td{margin:0;padding:0;"
            + "font:10px/12px '" + ManagedTextShapingTestAssets.FamilyName + "'}"
            + "table{border-collapse:collapse;border-spacing:0;border:0}td{border:0}"
            + "#running{position:running(heading)}"
            + "@page{size:800px 900px;margin:300px 20px 40px;"
            + "@top-center{content:" + (margin ? "element(heading)" : "none") + "}}</style>" + running + body;

        HtmlRenderOptions options = RunningMarginSizingOptions(budget, configuredProvider);
        // Capturing the same source and laying out all body pages fits the limit.
        // Only the newly repeated intrinsic traversal should exhaust this ledger.
        HtmlRenderDocument captureOnly = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(Source(false)), options);
        Assert.Equal(20, captureOnly.Pages.Count);
        HtmlDomLimitException error = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderEngine.Render(HtmlConversionDocument.Parse(Source(true)), options));
        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, error.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), error.LimitSource);
        Assert.Contains("running intrinsic text measurement", error.Message, StringComparison.Ordinal);

        options.MaxLayoutOperations = 1_000_000;
        HtmlRenderDocument complete = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(Source(true)), options);
        Assert.Equal(20, complete.Pages.Count);
        Assert.All(complete.Pages, page => Assert.Single(page.Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "@page @top-center element(heading)"));
        Assert.Equal(20, complete.Text.Split(new[] { 'B' }, StringSplitOptions.None).Length - 1);
        complete.RequireNoLoss();
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void HtmlPageMarginSizing_WrappedTableUsesTheWholeColumnGridWithoutClippingItsLastCell(
        bool cssDisplay, bool nestedFlex) {
        string table = RunningMarginSizingTable(cssDisplay);
        if (nestedFlex) table = "<div style='display:flex'><div>" + table + "</div></div>";
        string source = "<style>" + RunningMarginSizingTableStyles + "#running{overflow:hidden}</style>"
            + "<div id='running' style='position:running(heading)'>" + table
            + "</div><p>BODY</p>";
        HtmlRenderDocument rendered = RenderMarginSizing(900D,
            "@top-left{content:element(heading)}@top-center{content:'CENTRE TITLE'}@top-right{content:'R'}", body: source);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderSemanticGroup running = Assert.Single(page.Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "@page @top-left element(heading)");
        HtmlRenderText[] cells = EnumerateRenderVisuals(running.Visuals).OfType<HtmlRenderText>().ToArray();

        Assert.Equal(6300D / 17D, running.Width, 6);
        Assert.Equal(2700D / 17D, PageMarginTypographyText(page, "top-center").Width, 6);
        Assert.Equal(RunningMarginSizingColumns, cells.Select(cell => cell.Text).ToArray());
        Assert.All(cells, cell => Assert.InRange(cell.X + cell.Width, running.X, running.X + running.Width));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("input-button")]
    [InlineData("button")]
    [InlineData("select")]
    [InlineData("textarea")]
    [InlineData("progress")]
    [InlineData("meter")]
    public void HtmlPageMarginSizing_RootControlsUseTheSameIntrinsicOwnerAsWrappedControls(string kind) {
        const string label = "HEADER LABEL REQUIRES ROOM";
        string Control(bool running) {
            string attributes = " id='control' style=\"" + (running ? "position:running(heading);" : "")
                + "font:20px/20px '" + ManagedTextShapingTestAssets.FamilyName + "';padding:0;border:1px solid black\"";
            return kind switch {
                "input-button" => "<input type='button' value='" + label + "'" + attributes + ">",
                "button" => "<button" + attributes + ">" + label + "</button>",
                "select" => "<select" + attributes + "><option selected>" + label + "</option></select>",
                "textarea" => "<textarea cols='24' rows='2'" + attributes + ">" + label + "</textarea>",
                "progress" => "<progress value='1' max='2'" + attributes + "></progress>",
                "meter" => "<meter value='1' max='2'" + attributes + "></meter>",
                _ => throw new ArgumentOutOfRangeException(nameof(kind))
            };
        }
        const string boxes = "@top-left{content:element(heading)}@top-center{content:'CENTRE TITLE'}@top-right{content:'R'}";
        HtmlRenderDocument root = RenderMarginSizing(900D, boxes, body: Control(true) + "<p>BODY</p>");
        HtmlRenderDocument wrapped = RenderMarginSizing(900D, boxes,
            body: "<div style='position:running(heading)'>" + Control(false) + "</div><p>BODY</p>");
        HtmlRenderSemanticGroup RootGroup(HtmlRenderDocument document) => Assert.Single(
            Assert.Single(document.Pages).Scene.OfType<HtmlRenderSemanticGroup>(),
            group => group.Source == "@page @top-left element(heading)");
        HtmlRenderSemanticGroup directGroup = RootGroup(root), wrappedGroup = RootGroup(wrapped);
        double ControlWidth(HtmlRenderSemanticGroup group) => EnumerateRenderVisuals(group.Visuals)
            .Where(visual => visual.Source == (kind == "input-button" ? "input" : kind) + "#control"
                && visual is HtmlRenderShape or HtmlRenderFormField)
            .Max(visual => visual.Width);

        Assert.Equal(wrappedGroup.Width, directGroup.Width, 6);
        Assert.Equal(ControlWidth(wrappedGroup), ControlWidth(directGroup), 6);
        Assert.InRange(ControlWidth(directGroup), 1D, directGroup.Width);
        foreach (HtmlRenderDocument rendered in new[] { root, wrapped }) {
            if (kind is "select" or "textarea") {
                // The synthetic face cannot become an interactive PDF widget.
                // Its existing static fallback is reported at capture and relayout.
                HtmlDiagnostic[] losses = rendered.Diagnostics
                    .Where(diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None).ToArray();
                Assert.Equal(2, losses.Length);
                Assert.All(losses, diagnostic => {
                    Assert.Equal(HtmlRenderDiagnosticCodes.FormFieldTypographyStaticFallback, diagnostic.Code);
                    Assert.Equal(kind + "#control", diagnostic.Source);
                    Assert.Equal(HtmlDiagnosticSeverity.Warning, diagnostic.Severity);
                    Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
                });
            } else {
                rendered.RequireNoLoss();
            }
        }
    }

    [Theory]
    [InlineData(false, "flex")]
    [InlineData(true, "flex")]
    [InlineData(false, "grid")]
    [InlineData(true, "grid")]
    public void HtmlPageMarginSizing_SharedTableIntrinsicsAlsoSizeOrdinaryFlexAndGridWrappers(
        bool cssDisplay, string display) {
        string formatting = display == "grid"
            ? "#container{display:grid;grid-template-columns:max-content}"
            : "#container{display:flex}#measured{flex:0 0 auto}";
        string source = "<style>" + RunningMarginSizingTableStyles
            + formatting + "#measured{background:red}</style>"
            + "<div id='container'><div id='measured'>" + RunningMarginSizingTable(cssDisplay) + "</div></div>";
        HtmlRenderDocument rendered = RenderMarginSizing(900D, "", body: source);

        Assert.Equal(280D, NaturalLineShape(rendered, "div#measured").Width, 6);
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("table", false)]
    [InlineData("inline-table", false)]
    [InlineData("table", true)]
    [InlineData("inline-table", true)]
    public void HtmlPageMarginSizing_TableDisplayInputKeepsItsNativeControlSizing(string display, bool wrapped) {
        const string label = "HEADER LABEL REQUIRES ROOM";
        string Control(bool running) => "<input id='control' type='button' value='" + label
            + "' style=\"display:" + display + ";" + (running ? "position:running(heading);" : "")
            + "font:20px/20px '" + ManagedTextShapingTestAssets.FamilyName + "';padding:0;border:1px solid black\">";
        string running = wrapped
            ? "<div style='position:running(heading);overflow:hidden'>" + Control(false) + "</div>"
            : Control(true);
        HtmlRenderDocument rendered = RenderMarginSizing(900D,
            "@top-left{content:element(heading)}@top-center{content:'CENTRE TITLE'}@top-right{content:'R'}",
            body: running + "<p>BODY</p>");
        HtmlRenderDocument ordinary = RenderMarginSizing(900D, "", body: Control(false));
        HtmlRenderSemanticGroup group = RunningMarginSizingGroup(rendered);
        HtmlRenderShape border = NaturalLineShape(rendered, "input#control");
        HtmlRenderText text = Assert.Single(EnumerateRenderVisuals(group.Visuals).OfType<HtmlRenderText>());

        Assert.Equal(NaturalLineShape(ordinary, "input#control").Width, border.Width, 6);
        Assert.Equal(label.Length * 10D + 14D, border.Width, 6);
        Assert.InRange(border.X, group.X, group.X + group.Width);
        Assert.InRange(border.X + border.Width, group.X, group.X + group.Width);
        Assert.Equal(label, text.Text);
        Assert.Equal(20D, text.Font.Size, 6);
        Assert.True(text.Width >= label.Length * 10D);
        Assert.InRange(text.X, border.X, border.X + border.Width);
        Assert.InRange(text.X + text.Width, border.X, border.X + border.Width);
        rendered.RequireNoLoss();
        ordinary.RequireNoLoss();
    }

    [Theory]
    [InlineData("table", false)]
    [InlineData("inline-table", false)]
    [InlineData("table", true)]
    [InlineData("inline-table", true)]
    public void HtmlPageMarginSizing_TableDisplayImageKeepsItsNaturalDimensions(string display, bool wrapped) {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(180, 24));
        string Image(string imageDisplay, bool running) => "<img id='image' alt='Header mark' src='data:image/png;base64,"
            + data + "' style='display:" + imageDisplay + ";" + (running ? "position:running(heading)" : "") + "'>";
        string Running(string imageDisplay) => wrapped
            ? "<div style='position:running(heading);overflow:hidden'>" + Image(imageDisplay, false) + "</div>"
            : Image(imageDisplay, true);
        const string boxes = "@top-left{content:element(heading)}@top-center{content:'CENTRE TITLE'}@top-right{content:'R'}";
        HtmlRenderDocument rendered = RenderMarginSizing(900D, boxes, body: Running(display) + "<p>BODY</p>");
        HtmlRenderDocument ordinaryDisplay = RenderMarginSizing(900D, boxes, body: Running("block") + "<p>BODY</p>");
        HtmlRenderSemanticGroup group = RunningMarginSizingGroup(rendered);
        HtmlRenderImage image = Assert.Single(EnumerateRenderVisuals(group.Visuals).OfType<HtmlRenderImage>());

        Assert.Equal(RunningMarginSizingGroup(ordinaryDisplay).Width, group.Width, 6);
        Assert.True(group.Width >= 180D);
        Assert.Equal(180D, image.Width, 6);
        Assert.Equal(24D, image.Height, 6);
        Assert.InRange(image.X, group.X, group.X + group.Width);
        Assert.InRange(image.X + image.Width, group.X, group.X + group.Width);
        rendered.RequireNoLoss();
        ordinaryDisplay.RequireNoLoss();
    }

    [Theory]
    [InlineData("table")]
    [InlineData("inline-table")]
    public void HtmlPageMarginSizing_TableDisplayDoesNotChangeAnOrdinaryFlexControlMinimum(string display) {
        string Content(string controlDisplay) => "<div style='display:flex;width:120px'>"
            + "<input id='control' type='button' value='HEADER LABEL REQUIRES ROOM' style=\"display:" + controlDisplay
            + ";flex:1 1 auto;font:20px/20px '" + ManagedTextShapingTestAssets.FamilyName
            + "';padding:0;border:1px solid black\"><div style='width:40px;flex-shrink:0'></div></div>";
        HtmlRenderDocument rendered = RenderMarginSizing(900D, "", body: Content(display));
        HtmlRenderDocument ordinaryDisplay = RenderMarginSizing(900D, "", body: Content("inline-block"));

        Assert.Equal(NaturalLineShape(ordinaryDisplay, "input#control").Width,
            NaturalLineShape(rendered, "input#control").Width, 6);
        rendered.RequireNoLoss();
        ordinaryDisplay.RequireNoLoss();
    }

    private static HtmlRenderSemanticGroup RunningMarginSizingGroup(HtmlRenderDocument rendered) => Assert.Single(
        Assert.Single(rendered.Pages).Scene.OfType<HtmlRenderSemanticGroup>(),
        group => group.Source == "@page @top-left element(heading)");

    private static readonly string[] RunningMarginSizingColumns = { "COLAAAA", "COLBBBB", "COLCCCC", "COLDDDD" };
    private const string RunningMarginSizingTableStyles =
        "table,.sizing-table{display:table;border-collapse:collapse;border-spacing:0;border:0;padding:0}"
        + "tr,.sizing-row{display:table-row}td,.sizing-cell{display:table-cell;border:0;padding:0;white-space:nowrap}";

    private static string RunningMarginSizingTable(bool cssDisplay) => cssDisplay
        ? "<div class='sizing-table'><div class='sizing-row'>"
            + string.Concat(RunningMarginSizingColumns.Select(value => "<div class='sizing-cell'>" + value + "</div>")) + "</div></div>"
        : "<table><tbody><tr>" + string.Concat(RunningMarginSizingColumns.Select(value => "<td>" + value + "</td>")) + "</tr></tbody></table>";

    private static HtmlRenderOptions RunningMarginSizingOptions(int budget, bool configuredProvider) {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true,
            ViewportWidth = 800D, ViewportHeight = 900D,
            AllowSystemFontFallback = false, DefaultFontFamily = ManagedTextShapingTestAssets.FamilyName,
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser, MaxLayoutOperations = budget,
            TextShapingProvider = configuredProvider ? OfficeManagedTextShapingProvider.Instance : null
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(
                Enumerable.Range(32, 95).ToDictionary(scalar => scalar, _ => (ushort)500)));
        return options;
    }
}
