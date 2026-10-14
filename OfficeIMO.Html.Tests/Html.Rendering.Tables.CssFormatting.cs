using OfficeIMO.Html;
using OfficeIMO.Drawing;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlCssTable_AdjacentCellsUseAnonymousRowAndIntrinsicTableWidth() {
        const string html = "<style>body{margin:0}.cell{display:table-cell}.icon{width:80px;background:blue}"
            + ".body{padding:10px;background:gray}.icon>div{width:39px;height:37px;background:white}"
            + ".body>div{width:300px;height:40px;background:red}</style>"
            + "<div id='topnews' style='width:600px'><div id='icon' class='cell icon'><div id='symbol'></div></div>"
            + "<div id='body' class='cell body'><div id='detail'></div></div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, CssTableOptions());
        HtmlRenderShape icon = CssTableShape(rendered, "div#icon");
        HtmlRenderShape body = CssTableShape(rendered, "div#body");
        HtmlRenderShape detail = CssTableShape(rendered, "div#detail");

        Assert.Equal(0D, icon.X, 3);
        Assert.Equal(80D, icon.Width, 3);
        Assert.Equal(icon.Y, body.Y, 3);
        Assert.Equal(80D, body.X, 3);
        Assert.Equal(320D, body.Width, 3);
        Assert.Equal(60D, body.Height, 3);
        Assert.Equal(90D, detail.X, 3);
        Assert.Equal(10D, detail.Y, 3);
        Assert.Equal(13D, CssTableShape(rendered, "div#symbol").Y, 3);
        Assert.DoesNotContain(EnumerateTablePaginationScene(rendered.Pages[0].Scene).OfType<HtmlRenderSemanticGroup>(),
            group => group.Role == HtmlRenderSemanticGroupRole.Table);
        Assert.False(rendered.HasLoss);
    }

    [Fact]
    public void HtmlCssTable_AnonymousBoxesResetParentGeometryAndPreserveSelectorsLinksAndSiblings() {
        const string html = "<style>body{margin:0}#outer>.cell{display:table-cell;background:blue;width:80px}"
            + "#outer>.cell+span{background:gray;width:100px}</style>"
            + "<div id='outer' style='margin:10px;padding:7px;border:2px solid green;width:600px;font-size:18px;line-height:22px'>"
            + "<div id='before' style='height:12px;background:red'></div>"
            + "<a id='first' class='cell' href='https://example.com/alert'>Alert</a>"
            + "<div class='cell' style='display:none;width:400px'>Hidden</div>"
            + "<span id='second' class='cell'>Body</span>"
            + "<div id='after' style='height:8px;background:red'></div></div><a href='#first'>Jump to alert</a>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, CssTableOptions());
        HtmlRenderShape first = CssTableShape(rendered, "a#first");
        HtmlRenderShape second = CssTableShape(rendered, "span#second");
        HtmlRenderShape before = CssTableShape(rendered, "div#before");
        HtmlRenderShape after = CssTableShape(rendered, "div#after");

        Assert.Equal(before.X, first.X, 3);
        Assert.Equal(before.Y + before.Height, first.Y, 3);
        Assert.Equal(first.Y, second.Y, 3);
        Assert.Equal(first.X + first.Width, second.X, 3);
        Assert.Equal(80D, first.Width, 3);
        Assert.Equal(100D, second.Width, 3);
        Assert.Equal(first.Y + first.Height, after.Y, 3);
        HtmlRenderText alert = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Alert");
        Assert.Equal(18D, alert.Font.Size, 3);
        Assert.Equal("https://example.com/alert", alert.LinkUri);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Hidden");
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderNamedDestination>(), destination => destination.Name == "first");
        Assert.False(rendered.HasLoss);
    }

    [Theory]
    [InlineData(HtmlRenderMode.Continuous, "left")]
    [InlineData(HtmlRenderMode.Continuous, "right")]
    [InlineData(HtmlRenderMode.Paged, "left")]
    [InlineData(HtmlRenderMode.Paged, "right")]
    public void HtmlCssTable_AnonymousTableUsesAvailableBandBesideFloat(HtmlRenderMode mode, string side) {
        string html = "<style>body{margin:0}.cell{display:table-cell;background:gray;width:80px}</style>"
            + "<div style='width:600px'><div id='float' style='float:" + side + ";width:80px;height:100px;background:green'></div>"
            + "<div id='first' class='cell'><div style='height:40px'></div></div>"
            + "<div id='second' class='cell'><div style='height:40px'></div></div></div>";
        HtmlRenderOptions options = CssTableOptions();
        options.Mode = mode;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderShape floating = CssTableShape(rendered, "div#float");
        HtmlRenderShape first = CssTableShape(rendered, "div#first");
        HtmlRenderShape second = CssTableShape(rendered, "div#second");

        Assert.Equal(floating.Y, first.Y, 3);
        Assert.Equal(first.Y, second.Y, 3);
        Assert.Equal(first.X + first.Width, second.X, 3);
        if (side == "left") Assert.Equal(floating.X + floating.Width, first.X, 3);
        else Assert.True(second.X + second.Width <= floating.X);
        Assert.False(rendered.HasLoss);
    }

    [Theory]
    [InlineData("table", "table-row-group", "table-row", "table-cell", false)]
    [InlineData("table", "table-row-group", "table-row", "table-cell", true)]
    public void HtmlCssTable_ExplicitGroupsReuseColumnsAndKeepDeclaredDataSemantics(string table, string group, string row, string cell, bool dataRoles) {
        string roles(string role) => dataRoles ? " role='" + role + "'" : string.Empty;
        string html = "<style>body{margin:0}.table{display:" + table + ";width:240px;table-layout:fixed;border-spacing:0}"
            + ".group{display:" + group + "}.row{display:" + row + "}.cell{display:" + cell + ";padding:0;background:gray}</style>"
            + "<div id='css' class='table'" + roles("table") + "><div class='group'>"
            + "<div class='row'" + roles("row") + "><div id='a' class='cell' style='width:80px'" + roles("columnheader") + ">A</div>"
            + "<div id='b' class='cell'" + roles("cell") + ">B</div></div>"
            + "<div class='row'" + roles("row") + "><div id='c' class='cell'" + roles("cell") + ">C</div>"
            + "<div id='d' class='cell'" + roles("cell") + "><table id='nested' style='width:100px;margin:0;border-spacing:0'>"
            + "<tr><td style='padding:0'>Nested</td></tr></table></div></div></div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, CssTableOptions());
        Assert.Equal(80D, CssTableShape(rendered, "div#a").Width, 3);
        Assert.Equal(160D, CssTableShape(rendered, "div#b").Width, 3);
        Assert.Equal(CssTableShape(rendered, "div#a").X, CssTableShape(rendered, "div#c").X, 3);
        Assert.Equal(CssTableShape(rendered, "div#b").X, CssTableShape(rendered, "div#d").X, 3);
        HtmlRenderSemanticGroup[] tables = EnumerateTablePaginationScene(rendered.Pages[0].Scene)
            .OfType<HtmlRenderSemanticGroup>().Where(item => item.Role == HtmlRenderSemanticGroupRole.Table).ToArray();
        Assert.Equal(dataRoles ? 2 : 1, tables.Length);
        Assert.Contains(tables, item => item.Source == "table#nested");
        Assert.False(rendered.HasLoss);
    }

    [Fact]
    public void HtmlCssTable_OrdinaryRowContentUsesAnonymousCellsWithoutDroppingTextOrLinks() {
        const string html = "<body style='margin:0'><div style='display:table;border-spacing:0'>"
            + "<a style='display:table-row' href='https://example.com/row'>Before"
            + "<span id='middle' style='display:table-cell;padding:10px;background:gray'>Middle</span>After</a>"
            + "</div></body>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, CssTableOptions());
        HtmlRenderText[] text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Equal(new[] { "Before", "Middle", "After" }, text.Select(item => item.Text));
        Assert.All(text, item => Assert.Equal("https://example.com/row", item.LinkUri));
        Assert.True(text[1].X > text[0].X);
        Assert.True(text[2].X > text[1].X);
        Assert.False(rendered.HasLoss);
    }

    [Fact]
    public void HtmlCssTable_InlineTableIsAnAtomicBoxBetweenText() {
        const string html = "<body style='margin:0;font-size:12px;line-height:20px'>Before "
            + "<span style='display:inline-table;border-spacing:0'><span style='display:table-row'>"
            + "<span id='inline-first' style='display:table-cell;width:40px;background:gray'>One</span>"
            + "<span id='inline-second' style='display:table-cell;width:40px;background:blue'>Two</span>"
            + "</span></span> After</body>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, CssTableOptions());
        HtmlRenderText before = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Trim() == "Before");
        HtmlRenderText after = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Trim() == "After");
        HtmlRenderShape first = CssTableShape(rendered, "span#inline-first");
        HtmlRenderShape second = CssTableShape(rendered, "span#inline-second");
        Assert.Equal(first.Y, second.Y, 3);
        Assert.True(first.X >= before.X + before.TextAdvanceWidth);
        Assert.True(after.X >= second.X + second.Width);
        Assert.Equal(before.Y, after.Y, 3);
        Assert.False(rendered.HasLoss);
    }

    [Fact]
    public void HtmlCssTable_HeadersAndFootersRepeatWithFormattingOwnershipAndGenericSemantics() {
        string rows = string.Concat(Enumerable.Range(0, 14).Select(index => "<div class='row'><div class='cell'>Body"
            + index.ToString("D2") + "</div><div class='cell'>Value</div></div>"));
        string html = "<style>@page{size:200px 100px;margin:0}body{margin:0}.table{display:table;width:160px;border-spacing:0}"
            + ".header{display:table-header-group}.footer{display:table-footer-group}.row{display:table-row}"
            + ".cell{display:table-cell;font-size:10px;line-height:16px;padding:2px}</style>"
            + "<div class='table'><div class='header'><div class='row'><div class='cell'>Header</div><div class='cell'>Name</div></div></div>"
            + rows + "<div class='footer'><div class='row'><div class='cell'>Footer</div><div class='cell'>End</div></div></div></div>"
            + "<p style='margin:0'>After table</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlRenderPage[] bodyPages = rendered.Pages.Where(page => page.Visuals.OfType<HtmlRenderText>()
            .Any(text => text.Text.StartsWith("Body", StringComparison.Ordinal))).ToArray();
        Assert.True(bodyPages.Length >= 3);
        Assert.All(bodyPages, page => Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Header"));
        Assert.All(bodyPages, page => Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Footer"));
        HtmlRenderText[] allText = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
        Assert.Equal(14, allText.Count(text => text.Text.StartsWith("Body", StringComparison.Ordinal)));
        Assert.Equal(1, allText.Count(text => text.Text == "After table"));
        Assert.DoesNotContain(rendered.Pages.SelectMany(page => EnumerateTablePaginationScene(page.Scene)).OfType<HtmlRenderSemanticGroup>(),
            group => group.Role == HtmlRenderSemanticGroupRole.Table);
        Assert.False(rendered.HasLoss);
    }

    [Theory]
    [InlineData("<span><span style='display:table-cell'>Retained</span></span>")]
    [InlineData("<div style='display:table'><div style='display:table-column;width:80px'></div><div style='display:table-row'><div style='display:table-cell'>Retained</div></div></div>")]
    [InlineData("<div style='display:table'><div style='display:contents'><div style='display:table-cell'>Retained</div></div></div>")]
    public void HtmlCssTable_UnsupportedStructuresReportTypedLossAndRetainContent(string body) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render("<body style='margin:0'>" + body + "</body>", CssTableOptions());
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text == "Retained");
        Assert.True(rendered.HasLoss);
    }

    [Theory]
    [InlineData("table", false)]
    [InlineData("table-row", false)]
    [InlineData("table-row-group", false)]
    [InlineData("table", true)]
    [InlineData("table-row", true)]
    [InlineData("table-row-group", true)]
    public void HtmlCssTable_GeneratedContentOnFormattingBoxesReportsOmission(string display, bool native) {
        string table = native ? "table" : "div";
        string group = native ? "tbody" : "div";
        string row = native ? "tr" : "div";
        string cell = native ? "td" : "div";
        string target(string kind) => display == kind ? " id='target'" : string.Empty;
        string html = "<style>body{margin:0}#target::before{content:'GeneratedBefore'}#target::after{content:'GeneratedAfter'}</style>"
            + "<" + table + target("table") + " style='display:table;border-spacing:0'><" + group + target("table-row-group")
            + " style='display:table-row-group'><" + row + target("table-row") + " style='display:table-row'><" + cell
            + " style='display:table-cell'>Retained</" + cell + "></" + row + "></" + group + "></" + table + ">";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, CssTableOptions());
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported
            && diagnostic.LossKind == OfficeConversionLossKind.Omission && diagnostic.Source == (display == "table" ? table : display == "table-row" ? row : group) + "#target::before");
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported
            && diagnostic.LossKind == OfficeConversionLossKind.Omission && diagnostic.Source!.EndsWith("#target::after", StringComparison.Ordinal));
        Assert.DoesNotContain(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text.StartsWith("Generated", StringComparison.Ordinal));
        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Text == "Retained");
        Assert.True(rendered.HasLoss);
    }

    [Theory]
    [InlineData("content:none")]
    [InlineData("content:normal")]
    [InlineData("content:''")]
    [InlineData("content:'Suppressed';display:none")]
    public void HtmlCssTable_IneffectiveGeneratedContentDoesNotReportOmission(string declarations) {
        string html = "<style>body{margin:0}.box::before,.box::after{" + declarations + "}</style>"
            + "<div class='box' style='display:table'><div class='box' style='display:table-row-group'>"
            + "<div class='box' style='display:table-row'><div style='display:table-cell'>Retained</div></div></div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, CssTableOptions());
        Assert.False(rendered.HasLoss);
    }

    [Fact]
    public void HtmlCssTable_CellAndCaptionGeneratedContentRemainsSupported() {
        const string html = "<style>body{margin:0}#caption::before{content:'CaptionBefore'}#caption::after{content:'CaptionAfter'}"
            + "#cell::before{content:'CellBefore'}#cell::after{content:'CellAfter'}</style>"
            + "<div style='display:table'><div id='caption' style='display:table-caption'>Title</div>"
            + "<div style='display:table-row'><div id='cell' style='display:table-cell'>Value</div></div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, CssTableOptions());
        string text = rendered.Text;
        Assert.Contains("CaptionBefore", text);
        Assert.Contains("CaptionAfter", text);
        Assert.Contains("CellBefore", text);
        Assert.Contains("CellAfter", text);
        Assert.False(rendered.HasLoss);
    }

    [Theory]
    [InlineData("img", true)]
    [InlineData("input", true)]
    [InlineData("input-image", true)]
    [InlineData("img", false)]
    [InlineData("input", false)]
    [InlineData("input-image", false)]
    public void HtmlCssTable_ReplacedCellRetainsItsOwnImageOrControlAndInsets(string kind, bool authoredWidth) {
        const string image = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl4xd4AAAAASUVORK5CYII=";
        string tag = kind == "img" ? "img" : "input";
        string attributes = kind == "input" ? " value='CellValue'" : " src='" + image + "'";
        if (kind == "input-image") attributes += " type='image'";
        string element(string display) => "<" + tag + " id='replaced'" + attributes
            + " style='display:" + display + ";" + (authoredWidth ? "width:100px;" : string.Empty)
            + "height:30px;box-sizing:border-box;padding:3px;border:2px solid green;line-height:12px;vertical-align:top'>";
        string prefix = "<body style='margin:0;font-size:12px;line-height:12px'>";
        HtmlRenderDocument ordinary = HtmlRenderTestDriver.Render(prefix + element("inline-block"), CssTableOptions());
        HtmlRenderDocument table = HtmlRenderTestDriver.Render(prefix + "<div style='display:table;border-spacing:0'>"
            + element("table-cell") + "<div style='display:table-cell'>Sibling</div></div>", CssTableOptions());
        if (kind == "input") {
            HtmlRenderText expected = Assert.Single(ordinary.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "CellValue");
            HtmlRenderText actual = Assert.Single(table.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "CellValue");
            Assert.Equal(expected.X, actual.X, 3);
            Assert.Equal(expected.Y, actual.Y, 3);
            Assert.Equal(ordinary.Pages[0].Visuals.OfType<HtmlRenderShape>().Count(shape => shape.Source == "input#replaced"),
                table.Pages[0].Visuals.OfType<HtmlRenderShape>().Count(shape => shape.Source == "input#replaced"));
        } else {
            HtmlRenderImage expected = Assert.Single(ordinary.Pages[0].Visuals.OfType<HtmlRenderImage>());
            HtmlRenderImage actual = Assert.Single(table.Pages[0].Visuals.OfType<HtmlRenderImage>());
            Assert.Equal(expected.X, actual.X, 3);
            Assert.Equal(expected.Y, actual.Y, 3);
            Assert.Equal(expected.Width, actual.Width, 3);
            Assert.Equal(expected.Height, actual.Height, 3);
        }
        Assert.False(table.HasLoss);
    }

    [Theory]
    [InlineData("img")]
    [InlineData("input")]
    [InlineData("input-image")]
    public void HtmlCssTable_ReplacedCellUsesOnlyWinningCollapsedBordersAndOneOutline(string kind) {
        const string image = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl4xd4AAAAASUVORK5CYII=";
        string tag = kind == "img" ? "img" : "input";
        string attributes = kind == "input" ? " value='CellValue'" : " src='" + image + "'";
        if (kind == "input-image") attributes += " type='image'";
        string html = "<body style='margin:0;font-size:12px;line-height:12px'><div id='table' style='display:table;width:200px;border-collapse:collapse'>"
            + "<div style='display:table-row;border:4px solid purple'><" + tag + " id='replaced'" + attributes
            + " style='display:table-cell;width:100px;height:30px;box-sizing:border-box;padding:3px;background:white;border:2px solid green;outline:1px solid red;vertical-align:top'>"
            + "<div style='display:table-cell;border-left:6px solid blue'>Sibling</div></div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, CssTableOptions());
        HtmlRenderShape[] shapes = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>().ToArray();
        Assert.DoesNotContain(shapes, shape => shape.Shape.StrokeColor == OfficeColor.FromRgb(0, 128, 0));
        HtmlRenderShape shared = Assert.Single(shapes, shape => shape.Source == "div#table:collapsed-border-v-1-0");
        Assert.Equal(OfficeColor.Blue, shared.Shape.StrokeColor);
        Assert.Equal(6D, shared.Shape.StrokeWidth, 3);
        Assert.Single(shapes, shape => shape.Source == tag + "#replaced:outline");
        Assert.Contains(shapes, shape => shape.Source == tag + "#replaced" && shape.Shape.FillColor == OfficeColor.White);
        if (kind == "input") Assert.Equal(5D, Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "CellValue").X, 3);
        else Assert.Equal(5D, Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>()).X, 3);
        Assert.False(rendered.HasLoss);
    }

    [Theory]
    [InlineData("240px", 240D)]
    [InlineData("auto", 160D)]
    public void HtmlCssTable_NestedInlineTableContributesItsAtomicColumnWidths(string width, double expectedWidth) {
        string html = "<body style='margin:0;font-size:12px;line-height:20px'><div style='display:table;border-spacing:0'>"
            + "<div id='outer-cell' style='display:table-cell;background:gray'><span id='inner' style='display:inline-table;width:"
            + width + ";border-spacing:0;background:blue'><span style='display:table-row'>"
            + "<span style='display:table-cell;width:80px'>A</span><span style='display:table-cell;width:80px'>B</span></span></span></div>"
            + "<div id='next-cell' style='display:table-cell;background:red'>C</div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, CssTableOptions());
        HtmlRenderShape inner = CssTableShape(rendered, "span#inner");
        HtmlRenderShape outer = CssTableShape(rendered, "div#outer-cell");
        HtmlRenderText next = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "C");
        Assert.Equal(expectedWidth, inner.Width, 3);
        Assert.Equal(inner.Width, outer.Width, 3);
        Assert.True(next.X >= inner.X + inner.Width - 0.001D);
        Assert.False(rendered.HasLoss);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(false, true)]
    public void HtmlCssTable_AnonymousBoxesHaveDistinctStablePdfOwners(bool anonymousRows, bool paged) {
        string before = "Before" + (paged ? " " + string.Join(" ", Enumerable.Repeat("beforeword", 24)) : string.Empty);
        string after = "After" + (paged ? " " + string.Join(" ", Enumerable.Repeat("afterword", 24)) : string.Empty);
        string content = anonymousRows
            ? before + "<div role='row' style='display:table-row'><div role='cell' style='display:table-cell'>Middle</div></div>" + after
            : "<div role='row' style='display:table-row'>" + before + "<span role='cell' style='display:table-cell'>Middle</span>" + after + "</div>";
        string html = "<style>@page{size:200px 100px;margin:0}body{margin:0;font-size:10px;line-height:16px}</style>"
            + "<div role='table' style='display:table;width:180px;border-spacing:0'>" + content + "</div>";
        var options = new HtmlToPdfOptions(CssTableOptions()) { HonorCssPageRules = paged };
        PdfCore.PdfReadDocument pdf = PdfCore.PdfReadDocument.Open(HtmlConversionDocument.Parse(html).ToPdfBytes(options));
        Assert.Equal(anonymousRows ? 3 : 1, pdf.TaggedContent!.StructureElements.Count(element => element.StructureType == "TR"));
        Assert.Equal(3, pdf.TaggedContent.StructureElements.Count(element => element.StructureType == "TD"));
        Assert.Single(pdf.TaggedContent.StructureElements, element => element.StructureType == "Table");
        if (paged) Assert.True(pdf.Pages.Count > 1);
        string text = string.Concat(pdf.Pages.SelectMany(page => page.GetTextSpans()).Select(span => span.Text));
        Assert.True(text.IndexOf("Before", StringComparison.Ordinal) < text.IndexOf("Middle", StringComparison.Ordinal));
        Assert.True(text.IndexOf("Middle", StringComparison.Ordinal) < text.IndexOf("After", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(false, "header", false)]
    [InlineData(false, "footer", false)]
    [InlineData(true, "header", false)]
    [InlineData(true, "footer", false)]
    [InlineData(false, "header", true)]
    [InlineData(false, "footer", true)]
    [InlineData(true, "header", true)]
    [InlineData(true, "footer", true)]
    public void HtmlCssTable_OnlyFirstHeaderOrFooterGroupReordersAndRepeats(bool native, string kind, bool paged) {
        string table = native ? "table" : "div";
        string group = native ? kind == "header" ? "thead" : "tfoot" : "div";
        string row = native ? "tr" : "div";
        string cell = native ? "td" : "div";
        string Row(string value) => "<" + row + " style='display:table-row'><" + cell
            + " style='display:table-cell;padding:0'>" + value + "</" + cell + "></" + row + ">";
        string Group(string value) => "<" + group + " style='display:table-" + kind + "-group'>" + Row(value) + "</" + group + ">";
        string html = "<style>@page{size:200px 100px;margin:0}body{margin:0;font-size:10px;line-height:16px}</style>"
            + "<" + table + " style='display:table;width:180px;border-spacing:0'>" + Group("First") + Row("Body0")
            + Group("Second") + string.Concat(Enumerable.Range(1, paged ? 12 : 1).Select(index => Row("Body" + index))) + "</" + table + ">";
        HtmlRenderOptions options = CssTableOptions();
        options.Mode = paged ? HtmlRenderMode.Paged : HtmlRenderMode.Continuous;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        string[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(item => item.Text).ToArray();
        Assert.Equal(1, text.Count(value => value == "Second"));
        Assert.True(Array.IndexOf(text, "Body0") < Array.IndexOf(text, "Second"));
        Assert.True(Array.IndexOf(text, "Second") < Array.IndexOf(text, "Body1"));
        if (paged) Assert.True(text.Count(value => value == "First") > 1);
        else Assert.Equal(kind == "header" ? new[] { "First", "Body0", "Second", "Body1" } : new[] { "Body0", "Second", "Body1", "First" }, text);
        Assert.False(rendered.HasLoss);
    }

    [Theory]
    [InlineData(false, "header")]
    [InlineData(false, "footer")]
    [InlineData(true, "header")]
    [InlineData(true, "footer")]
    public void HtmlCssTable_EmptyFirstGroupKeepsLaterGroupInBodySourceOrder(bool native, string kind) {
        string table = native ? "table" : "div";
        string group = native ? kind == "header" ? "thead" : "tfoot" : "div";
        string row = native ? "tr" : "div";
        string cell = native ? "td" : "div";
        string Row(string value) => "<" + row + " style='display:table-row'><" + cell
            + " style='display:table-cell;padding:0'>" + value + "</" + cell + "></" + row + ">";
        string html = "<style>@page{size:200px 100px;margin:0}body{margin:0;font-size:10px;line-height:16px}</style>"
            + "<" + table + " style='display:table;width:180px;border-spacing:0'><" + group + " style='display:table-" + kind + "-group'></" + group + ">"
            + Row("Before") + "<" + group + " style='display:table-" + kind + "-group'>" + Row("Second") + "</" + group + ">"
            + string.Concat(Enumerable.Range(0, 12).Select(index => Row("After" + index))) + "</" + table + ">";
        HtmlRenderOptions options = CssTableOptions();
        options.Mode = HtmlRenderMode.Paged;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        string[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Select(item => item.Text).ToArray();
        Assert.True(rendered.Pages.Count > 1);
        Assert.Equal(1, text.Count(value => value == "Second"));
        Assert.Equal(new[] { "Before", "Second", "After0" }, text.Take(3));
        Assert.False(rendered.HasLoss);
    }

    private static HtmlRenderOptions CssTableOptions() => new HtmlRenderOptions {
        ViewportWidth = 700D,
        Margins = HtmlRenderMargins.All(0D)
    };

    private static HtmlRenderShape CssTableShape(HtmlRenderDocument document, string source) =>
        Assert.Single(document.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>(), shape => shape.Source == source);
}
