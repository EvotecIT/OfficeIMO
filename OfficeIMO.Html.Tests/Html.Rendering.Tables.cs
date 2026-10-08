using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlTables_WithoutAuthoredBordersDoNotPaintCellGrid() {
        const string html = "<body style='margin:0'>"
            + "<table><tr><td id='plain'>Plain</td></tr></table>"
            + "<table border='0'><tr><td id='zero'>Zero</td></tr></table>"
            + "<table><tr><td id='authored' style='border:1px solid red'>Authored</td></tr></table>"
            + "<table id='outer' style='border:2px solid red'><tr><td id='outer-cell'>Outer only</td></tr></table>"
            + "<table id='legacy' border='1'><tr><td id='legacy-cell'>Legacy grid</td></tr></table>"
            + "</body>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { ViewportWidth = 300D, Margins = HtmlRenderMargins.All(0D) });
        HtmlRenderShape[] shapes = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>().ToArray();

        Assert.DoesNotContain(shapes, shape => shape.Source is "td#plain" or "td#zero" or "td#outer-cell");
        Assert.Contains(shapes, shape => shape.Source == "td#authored");
        Assert.Contains(shapes, shape => shape.Source == "table#outer");
        Assert.Contains(shapes, shape => shape.Source == "table#legacy");
        Assert.Contains(shapes, shape => shape.Source == "td#legacy-cell");
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void HtmlTable_AutoLayoutKeepsAuthoredCellWidthWhenAnotherColumnCanGrow(bool nestedBreak, bool tabBeforeBreak) {
        string labels = (tabBeforeBreak ? "Home\t<br>" : "Home<br>") + "Reports<br>Document Queue<br>Administration";
        if (nestedBreak) labels = "<span>" + labels + "</span>";
        string html = "<body style='margin:0'><table style='width:320px;margin:0;border-spacing:2px'>"
            + "<tr><td id='nav' style='width:170px;padding:8px;background:blue'>" + labels + "</td>"
            + "<td id='content' style='background:white'>Content</td></tr></table></body>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) });
        HtmlRenderShape nav = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "td#nav" && shape.Shape.FillColor == OfficeColor.Blue);

        Assert.Equal(186D, nav.Width, 3);
    }

    [Fact]
    public void HtmlMarquee_UsesBlockWidthBeforeFollowingHeading() {
        const string html = "<body style='margin:0'><marquee id='notice' style='background:yellow'>Notice</marquee>"
            + "<h1 style='margin:0'>Heading</h1></body>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { ViewportWidth = 320D, Margins = HtmlRenderMargins.All(0D) });

        HtmlRenderShape notice = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "marquee#notice" && shape.Shape.FillColor == OfficeColor.Yellow);
        HtmlRenderText heading = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.Text == "Heading");
        Assert.Equal(320D, notice.Width, 3);
        Assert.True(heading.Y >= notice.Y + notice.Height);
    }

    [Fact]
    public void HtmlTable_PreservedTabBeforeNestedBreakContributesItsTabStopWidth() {
        const string html = "<body style='margin:0'><table style='margin:0;border-spacing:0'>"
            + "<tr><td id='tabbed' style='padding:0;background:blue;white-space:pre;tab-size:32px'><span>A\tB<br>C</span></td></tr>"
            + "</table></body>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { ViewportWidth = 320D, Margins = HtmlRenderMargins.All(0D) });
        HtmlRenderShape tabbed = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "td#tabbed" && shape.Shape.FillColor == OfficeColor.Blue);

        Assert.True(tabbed.Width >= 38D);
    }

    [Fact]
    public void HtmlTables_ApplyBrowserCaptionAndHeaderDefaultsWithoutOverridingAuthoredStyles() {
        const string prefix = "<body style='margin:0'><table style='width:240px;margin:0'>";
        const string suffix = "</table></body>";
        HtmlRenderDocument defaults = HtmlRenderTestDriver.Render(prefix
            + "<caption>Report caption</caption><tr><th>Region heading</th></tr>" + suffix,
            new HtmlRenderOptions { ViewportWidth = 300D, Margins = HtmlRenderMargins.All(0D) });
        HtmlRenderDocument authored = HtmlRenderTestDriver.Render(prefix
            + "<caption style='text-align:left'>Report caption</caption><tr><th style='font-weight:normal'>Region heading</th></tr>" + suffix,
            new HtmlRenderOptions { ViewportWidth = 300D, Margins = HtmlRenderMargins.All(0D) });

        HtmlRenderText defaultCaption = Assert.Single(defaults.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Report caption");
        HtmlRenderText authoredCaption = Assert.Single(authored.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Report caption");
        HtmlRenderText defaultHeader = Assert.Single(defaults.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Region heading");
        HtmlRenderText authoredHeader = Assert.Single(authored.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Region heading");

        Assert.True(defaultCaption.X > authoredCaption.X + 20D);
        Assert.True((defaultHeader.Font.Style & OfficeFontStyle.Bold) != 0);
        Assert.True((authoredHeader.Font.Style & OfficeFontStyle.Bold) == 0);
    }

    [Fact]
    public void HtmlTableCentersDeclaredBorderBoxAndUsesHeaderDefaultUnderInheritedNormalWeight() {
        const string html = "<body style='margin:0;font-weight:normal'>"
            + "<table id='report' style='width:240px;box-sizing:border-box;padding:10px;margin:0 auto;background:red'>"
            + "<tr><th>Header</th></tr></table></body>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { ViewportWidth = 300D, Margins = HtmlRenderMargins.All(0D) });

        HtmlRenderShape table = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "table#report" && shape.Shape.FillColor == OfficeColor.Red);
        HtmlRenderText header = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Header");
        Assert.Equal(30D, table.X, 3);
        Assert.Equal(240D, table.Width, 3);
        Assert.True((header.Font.Style & OfficeFontStyle.Bold) != 0);
    }

    [Fact]
    public void HtmlTables_LegacyCenterAlignUsesAutoMarginsUnlessCssOverridesThem() {
        const string prefix = "<style>body{margin:0}main{width:500px}</style><main>";
        const string suffix = "<tr><td>Cell</td></tr></table></main>";
        var options = new HtmlRenderOptions { ViewportWidth = 500D, Margins = HtmlRenderMargins.All(0D) };

        HtmlRenderDocument auto = HtmlRenderTestDriver.Render(prefix + "<table align='center'>" + suffix, options);
        HtmlRenderDocument fixedWidth = HtmlRenderTestDriver.Render(prefix + "<table align='center' style='width:160px'>" + suffix, options);
        HtmlRenderDocument cssOverride = HtmlRenderTestDriver.Render(prefix + "<table align='center' style='width:160px;margin:0'>" + suffix, options);

        HtmlRenderSemanticGroup autoTable = Assert.Single(EnumerateTablePaginationScene(auto.Pages[0].Scene)
            .OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Table);
        HtmlRenderSemanticGroup fixedTable = Assert.Single(EnumerateTablePaginationScene(fixedWidth.Pages[0].Scene)
            .OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Table);
        HtmlRenderSemanticGroup overriddenTable = Assert.Single(EnumerateTablePaginationScene(cssOverride.Pages[0].Scene)
            .OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Table);

        Assert.Equal((500D - autoTable.Width) / 2D, autoTable.X, 1);
        Assert.Equal(170D, fixedTable.X, 1);
        Assert.Equal(0D, overriddenTable.X, 1);
    }

    [Theory]
    [InlineData("width='200'")]
    [InlineData("style='width:200px;border-spacing:0'")]
    public void HtmlTables_ExplicitNestedTableDoesNotFlattenItsRowsIntoOuterPreferredWidth(string widthAttribute) {
        string html = "<body style='margin:0'><main style='width:500px'>"
            + "<table align='center' style='border-spacing:1px'><tr><td style='padding:0'>"
            + "<table " + widthAttribute + "><tr><td>Alpha Beta Gamma Delta Epsilon Zeta Eta Theta Iota Kappa</td></tr>"
            + "<tr><td>Lambda Mu Nu Xi Omicron Pi Rho Sigma Tau Upsilon</td></tr></table>"
            + "</td></tr></table></main></body>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { ViewportWidth = 500D, Margins = HtmlRenderMargins.All(0D) });
        HtmlRenderSemanticGroup[] tables = EnumerateTablePaginationScene(rendered.Pages[0].Scene)
            .OfType<HtmlRenderSemanticGroup>().Where(group => group.Role == HtmlRenderSemanticGroupRole.Table).ToArray();

        Assert.Equal(2, tables.Length);
        HtmlRenderSemanticGroup outer = tables.OrderByDescending(table => table.Width).First();
        Assert.InRange(outer.Width, 200D, 230D);
        Assert.Equal((500D - outer.Width) / 2D, outer.X, 1);
    }

    [Fact]
    public void HtmlTables_SizedNestedTableSeparatesSurroundingTextDuringOuterSizing() {
        const string html = "<body style='margin:0'><main style='width:500px'>"
            + "<table align='center' style='border-spacing:1px'><tr><td style='padding:0'>"
            + "Before alpha beta gamma<table width='200'><tr><td>Inner</td></tr></table>After delta epsilon zeta"
            + "</td></tr></table></main></body>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { ViewportWidth = 500D, Margins = HtmlRenderMargins.All(0D) });
        HtmlRenderSemanticGroup outer = EnumerateTablePaginationScene(rendered.Pages[0].Scene)
            .OfType<HtmlRenderSemanticGroup>().Where(group => group.Role == HtmlRenderSemanticGroupRole.Table)
            .OrderByDescending(table => table.Width).First();

        Assert.InRange(outer.Width, 200D, 230D);
        Assert.Equal((500D - outer.Width) / 2D, outer.X, 1);
    }

    [Fact]
    public void HtmlTables_InlineDisplayedNestedTableStaysInOuterTextSizing() {
        const string html = "<body style='margin:0'><main style='width:500px'>"
            + "<table style='border-spacing:1px'><tr><td style='padding:0'>"
            + "Before alpha beta gamma<table style='display:inline;width:200px'><tr><td>Inner</td></tr></table>After delta epsilon zeta"
            + "</td></tr></table></main></body>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { ViewportWidth = 500D, Margins = HtmlRenderMargins.All(0D) });
        HtmlRenderSemanticGroup outer = EnumerateTablePaginationScene(rendered.Pages[0].Scene)
            .OfType<HtmlRenderSemanticGroup>().Where(group => group.Role == HtmlRenderSemanticGroupRole.Table)
            .OrderByDescending(table => table.Width).First();

        Assert.True(outer.Width > 250D);
    }

    [Fact]
    public void HtmlTables_RejectRowsAndColumnsBeforeAllocatingLayoutTracks() {
        var rowOptions = new HtmlRenderOptions { MaxTableRows = 1 };
        HtmlDomLimitException rowException = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render("<table><tr><td>A</td></tr><tr><td>B</td></tr></table>", rowOptions));
        Assert.Equal(HtmlRenderDiagnosticCodes.TableLimitExceeded, rowException.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxTableRows), rowException.LimitSource);

        var columnOptions = new HtmlRenderOptions { MaxTableColumns = 2 };
        HtmlDomLimitException columnException = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render("<table><col span='1000'><tr><td>A</td></tr></table>", columnOptions));
        Assert.Equal(HtmlRenderDiagnosticCodes.TableLimitExceeded, columnException.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxTableColumns), columnException.LimitSource);
    }

    [Fact]
    public void HtmlTables_BoundCollapsedBorderCandidateExpansion() {
        const string html = "<table style='border-collapse:collapse'><tr><td style='border:1px solid red'>A</td><td style='border:1px solid red'>B</td></tr></table>";
        var options = new HtmlRenderOptions { MaxCollapsedTableBorderSegments = 2 };

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() => HtmlRenderTestDriver.Render(html, options));

        Assert.Equal(HtmlRenderDiagnosticCodes.CollapsedTableBorderLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxCollapsedTableBorderSegments), exception.LimitSource);
    }

    [Fact]
    public void HtmlTableCell_StacksBlockDescendantsAndRetainsHeadingSemantics() {
        const string html = """
            <table style="width:320px;border-collapse:collapse">
              <tr><td style="padding:12px">
                <h1 style="margin:0 0 8px">Action Required</h1>
                <p style="margin:0 0 8px">Review the deployment package.</p>
                <div style="padding:8px;background:#fff4d6">Two checks need attention.</div>
              </td></tr>
            </table>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { ViewportWidth = 340D, Margins = HtmlRenderMargins.All(0D) });
        HtmlRenderText heading = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Action Required");
        HtmlRenderText paragraph = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Review the deployment", StringComparison.Ordinal));
        HtmlRenderText notice = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text.Contains("Two checks", StringComparison.Ordinal));

        Assert.True(heading.Y < paragraph.Y);
        Assert.True(paragraph.Y < notice.Y);
        HtmlRenderHeading outline = Assert.Single(rendered.Headings);
        Assert.Equal("Action Required", outline.Text);
        Assert.Equal(1, outline.Level);

        string pdfText = PdfCore.PdfReadDocument.Open(
            HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions())).ExtractText();
        Assert.Contains("Action", pdfText, StringComparison.Ordinal);
        Assert.Contains("Review", pdfText, StringComparison.Ordinal);
        Assert.Contains("Two", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlTable_PaginatesInsideOneOversizedCellAtLineBoundaries() {
        const string html = """
            <table style="width:100px;border-collapse:collapse"><tbody><tr><td style="font-size:12px;line-height:20px">
              One<br>Two<br>Three<br>Four<br>Five<br>Six
            </td></tr></tbody></table>
            <div id="after-table" style="height:20px;background:#00ff00">After</div>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(2D, 70D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.True(rendered.Pages.Count >= 2);
        Assert.All(rendered.Pages, page => Assert.True(page.Visuals.Count > 1));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Theory]
    [InlineData("top")]
    [InlineData("bottom")]
    public void HtmlTables_CaptionSidePaintsStyledCaptionAroundGridAcrossBackends(string side) {
        string html = "<body style='margin:0'><table id='table' style='width:80px;margin:0;caption-side:" + side + ";font-size:8px;line-height:10px'>"
            + "<caption id='caption' style='padding:2px;background:#ff0000'>CaptionPdf</caption>"
            + "<tr><td>CellPdf</td></tr></table></body>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderText caption = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "CaptionPdf");
        HtmlRenderText cell = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "CellPdf");
        HtmlRenderShape captionBackground = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "caption#caption" && shape.Shape.FillColor == OfficeColor.Red);
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlToPdfOptions pdfOptions = new HtmlToPdfOptions();
        pdfOptions = new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 50D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfBytes(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(80D, captionBackground.Width, 3);
        if (side == "top") Assert.True(caption.Y < cell.Y);
        else Assert.True(caption.Y > cell.Y);
        Assert.Contains("CaptionPdf", svg, StringComparison.Ordinal);
        Assert.Contains("CellPdf", svg, StringComparison.Ordinal);
        Assert.Contains("CaptionPdf", pdfText, StringComparison.Ordinal);
        Assert.Contains("CellPdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlTables_EmptyGridRetainsItsCaption() {
        const string html = "<table style='width:60px;margin:0'><caption id='caption'>CaptionOnly</caption></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Contains("CaptionOnly", string.Concat(rendered.Text.Where(character => !char.IsWhiteSpace(character))), StringComparison.Ordinal);
        HtmlRenderSemanticGroup table = Assert.Single(rendered.Pages[0].Scene.OfType<HtmlRenderSemanticGroup>());
        Assert.Equal(HtmlRenderSemanticGroupRole.Table, table.Role);
        Assert.Contains(table.Visuals.OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Caption);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.EmptyTable);
    }

    [Fact]
    public void HtmlTables_TopCaptionIsNotRepeatedWhenContinuationRelayoutChangesPageWidth() {
        const string html = """
            <style>
              @page { size:300px 70px; margin:0; }
              @page :left { size:180px 70px; }
              table { width:100%; margin:0; font-size:8px; line-height:12px; }
            </style>
            <table><caption>OneCaption</caption><tbody>
              <tr><td>Row one has enough text to participate in width-sensitive relayout.</td></tr>
              <tr><td>Row two has enough text to participate in width-sensitive relayout.</td></tr>
              <tr><td>Row three has enough text to participate in width-sensitive relayout.</td></tr>
              <tr><td>Row four has enough text to participate in width-sensitive relayout.</td></tr>
            </tbody></table>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.True(rendered.Pages.Count > 1);
        Assert.Single(rendered.Pages.SelectMany(page => page.Visuals.OfType<HtmlRenderText>()), text => text.Text == "OneCaption");
    }

    [Fact]
    public void HtmlTables_AutoLayoutAllocatesColumnsFromIntrinsicCellContent() {
        const string html = "<table style='width:100px;margin:0;table-layout:auto;font-size:8px;line-height:10px'><tr>"
            + "<td id='wide' style='background:red'>WWWWWWWWWW</td><td id='narrow' style='background:blue'>i</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape wide = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#wide" && shape.Shape.FillColor == OfficeColor.Red);
        HtmlRenderShape narrow = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#narrow" && shape.Shape.FillColor == OfficeColor.Blue);

        Assert.True(wide.Width > narrow.Width * 2D);
        Assert.Equal(100D, wide.Width + narrow.Width, 3);
        Assert.Equal(wide.Width, narrow.X, 3);
    }

    [Fact]
    public void HtmlTables_WidthAutoShrinksToPreferredColumnsWithinContainingBlock() {
        const string html = "<style>body{margin:0}main{width:596px}table{margin:0;border-collapse:collapse;font-size:16px;line-height:24px}"
            + "th,td{border:1px solid black;padding:8px 16px}</style>"
            + "<main><table id='grid'><tr><th>Contaminant</th><th>Secondary Standard</th></tr>"
            + "<tr><td>Total Dissolved Solids</td><td>500 mg/L</td></tr>"
            + "<tr><td>Odor</td><td>3 threshold odor number</td></tr></table></main>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 620D,
            ViewportHeight = 400D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderSemanticGroup table = Assert.Single(EnumerateTablePaginationScene(rendered.Pages[0].Scene)
            .OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Table);

        Assert.InRange(table.Width, 350D, 450D);
    }

    [Fact]
    public void HtmlTables_WidthAutoRespectsCaptionMinimumAndCentersAutoMargins() {
        const string html = "<style>body{margin:0}main{width:596px}table{margin:0 auto;border-collapse:collapse;font-size:20px}"
            + "td{border:1px solid black;padding:4px}</style><main><table>"
            + "<caption>UnbreakableCaptionMinimumWidth</caption><tr><td>A</td><td>B</td></tr></table></main>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 620D,
            ViewportHeight = 400D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderSemanticGroup table = Assert.Single(EnumerateTablePaginationScene(rendered.Pages[0].Scene)
            .OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Table);

        // Chromium's default-serif control uses a 294.7px caption minimum in 596px.
        Assert.InRange(table.Width, 275D, 330D);
        Assert.InRange(table.X, 133D, 161D);
        Assert.Equal((596D - table.Width) / 2D, table.X, 1);
    }

    [Fact]
    public void HtmlTables_WidthAutoRespectsNoWrapCaptionLine() {
        const string html = "<style>body{margin:0}main{width:596px}table{margin:0 auto;font-size:20px}"
            + "caption{white-space:nowrap}</style><main><table>"
            + "<caption>Alpha beta gamma delta</caption><tr><td>A</td></tr></table></main>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 620D,
            ViewportHeight = 400D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderSemanticGroup table = Assert.Single(EnumerateTablePaginationScene(rendered.Pages[0].Scene)
            .OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Table);

        Assert.InRange(table.Width, 190D, 350D);
        Assert.Equal((596D - table.Width) / 2D, table.X, 1);
    }

    [Fact]
    public void HtmlTables_WidthAutoMeasuresCaptionChildImageAndGeneratedCellImage() {
        string captionImage = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(120, 10));
        string cellImage = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(200, 10));
        string html = "<style>body{margin:0}main{width:596px}table{margin:0 auto}"
            + "caption span{font-size:40px}td::before{content:url('data:image/png;base64," + cellImage + "')}"
            + "</style><main><table><caption><span>Wide</span><img src='data:image/png;base64,"
            + captionImage + "'></caption><tr><td></td></tr></table></main>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 620D,
            ViewportHeight = 400D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderSemanticGroup table = Assert.Single(EnumerateTablePaginationScene(rendered.Pages[0].Scene)
            .OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Table);

        Assert.InRange(table.Width, 200D, 350D);
        Assert.Equal((596D - table.Width) / 2D, table.X, 1);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>(), image => image.Width == 200D);
    }

    [Fact]
    public void HtmlTables_WidthAutoMeasuresGeneratedCaptionImage() {
        string image = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(220, 10));
        string html = "<style>body{margin:0}main{width:596px}table{margin:0 auto}"
            + "caption::before{content:url('data:image/png;base64," + image + "')}"
            + "</style><main><table><caption></caption><tr><td>A</td></tr></table></main>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 620D,
            ViewportHeight = 400D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderSemanticGroup table = Assert.Single(EnumerateTablePaginationScene(rendered.Pages[0].Scene)
            .OfType<HtmlRenderSemanticGroup>(), group => group.Role == HtmlRenderSemanticGroupRole.Table);

        Assert.InRange(table.Width, 220D, 260D);
        Assert.Equal((596D - table.Width) / 2D, table.X, 1);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>(), visual => visual.Width == 220D);
    }

    [Fact]
    public void HtmlTables_AutoLayoutMeasuresQuotedTabTextWithEmbeddedFont() {
        string? installedFamily = new[] { "Trebuchet MS", "Arial", "Calibri", "Liberation Sans", "DejaVu Sans" }
            .FirstOrDefault(candidate => PdfCore.PdfEmbeddedFontFamily.TryFromSystem(candidate, out _));
        if (installedFamily == null) return;

        var options = new HtmlToPdfOptions();
        options.ResourcePolicy.AllowDocumentFontEmbedding = true;
        string html = "<table style=\"font-family:'" + installedFamily + "'\"><tr><td>\"A\tB\"</td><td>Next</td></tr></table>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(options);

        Assert.Contains("A B", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlTables_AutoLayoutIncludesIntrinsicReplacedImageWidth() {
        string imageData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(80, 10));
        string html = "<table style='width:100px;margin:0;table-layout:auto;font-size:8px;line-height:10px'><tr>"
            + "<td id='image-cell' style='background:red'><img src='data:image/png;base64," + imageData + "'></td>"
            + "<td id='text-cell' style='background:blue'>i</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape imageCell = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#image-cell" && shape.Shape.FillColor == OfficeColor.Red);
        HtmlRenderShape textCell = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#text-cell" && shape.Shape.FillColor == OfficeColor.Blue);

        Assert.True(imageCell.Width > textCell.Width * 4D);
        Assert.Equal(100D, imageCell.Width + textCell.Width, 3);
    }

    [Fact]
    public void HtmlTables_FixedLayoutHonorsColWidthsAndDistributesRemainder() {
        const string html = "<table style='width:100px;margin:0;table-layout:fixed;font-size:8px;line-height:10px'>"
            + "<colgroup><col style='width:70px'><col></colgroup><tr>"
            + "<td id='first' style='background:red'>A</td><td id='second' style='background:blue'>Long content does not resize this column</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape first = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#first" && shape.Shape.FillColor == OfficeColor.Red);
        HtmlRenderShape second = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#second" && shape.Shape.FillColor == OfficeColor.Blue);

        Assert.Equal(70D, first.Width, 3);
        Assert.Equal(30D, second.Width, 3);
        Assert.Equal(70D, second.X, 3);
    }

    [Fact]
    public void HtmlTables_SeparateBordersApplyHorizontalAndVerticalSpacing() {
        const string html = "<table style='width:100px;margin:0;table-layout:fixed;border-collapse:separate;border-spacing:4px 3px;font-size:8px;line-height:10px'>"
            + "<tr><td id='first' style='background:red'>A</td><td id='second' style='background:blue'>B</td></tr>"
            + "<tr><td id='third' style='background:lime'>C</td><td>D</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 50D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape first = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#first" && shape.Shape.FillColor == OfficeColor.Red);
        HtmlRenderShape second = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#second" && shape.Shape.FillColor == OfficeColor.Blue);
        HtmlRenderShape third = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#third" && shape.Shape.FillColor == OfficeColor.Lime);

        Assert.Equal((4D, 3D, 44D), (first.X, first.Y, first.Width));
        Assert.Equal((52D, 3D, 44D), (second.X, second.Y, second.Width));
        Assert.Equal(20D, third.Y, 3);
    }

    [Fact]
    public void HtmlTables_CollapsedBordersIgnoreBorderSpacingInGridGeometry() {
        const string html = "<table style='width:100px;margin:0;table-layout:fixed;border-collapse:collapse;border-spacing:10px;font-size:8px;line-height:10px'>"
            + "<tr><td id='first' style='background:red'>A</td><td id='second' style='background:blue'>B</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderShape first = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#first" && shape.Shape.FillColor == OfficeColor.Red);
        HtmlRenderShape second = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "td#second" && shape.Shape.FillColor == OfficeColor.Blue);

        Assert.Equal((0D, 0D, 50D), (first.X, first.Y, first.Width));
        Assert.Equal((50D, 0D, 50D), (second.X, second.Y, second.Width));
    }

    [Fact]
    public void HtmlTables_CollapsedBordersResolveSharedCellEdgeOnceAcrossBackends() {
        const string html = "<table id='conflict' style='width:100px;margin:0;table-layout:fixed;border-collapse:collapse;font-size:8px;line-height:10px'><tr>"
            + "<td style='border:1px solid black;border-right:5px solid red'>LeftPdf</td>"
            + "<td style='border:1px solid black;border-left:2px solid blue'>RightPdf</td></tr></table>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 110D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderShape shared = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "table#conflict:collapsed-border-v-1-0");
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        HtmlToPdfOptions pdfOptions = new HtmlToPdfOptions();
        pdfOptions = new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(110D / HtmlRenderOptions.CssPixelsPerInch, 30D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        };
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfBytes(pdfOptions);
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(pdf).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(OfficeColor.Red, shared.Shape.StrokeColor);
        Assert.Equal(5D, shared.Shape.StrokeWidth, 3);
        Assert.Contains("stroke=\"#FF0000\"", svg, StringComparison.Ordinal);
        Assert.Contains("LeftPdf", pdfText, StringComparison.Ordinal);
        Assert.Contains("RightPdf", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlTables_CollapsedThreeDimensionalBordersPreserveConflictOrderAndTwoTonePaint() {
        const string html = "<table id='three-d' style='width:80px;margin:4px;table-layout:fixed;border-collapse:collapse;font-size:8px;line-height:10px'><tr>"
            + "<td style='border:2px solid black;border-right:6px groove #808080'>LeftPdf</td>"
            + "<td style='border:2px solid black;border-left:6px ridge #808080'>RightPdf</td></tr></table>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 100D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        List<HtmlRenderShape> shared = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>()
            .Where(shape => shape.Source != null && shape.Source.StartsWith("table#three-d:collapsed-border-v-1-0", StringComparison.Ordinal))
            .ToList();
        OfficeColor dark = OfficeColorTransforms.Shade(OfficeColor.Gray, 0.55D);
        OfficeColor light = OfficeColorTransforms.Tint(OfficeColor.Gray, 0.55D);
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(100D / HtmlRenderOptions.CssPixelsPerInch, 40D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Equal(2, shared.Count);
        Assert.Contains(shared, shape => shape.Shape.StrokeColor == dark);
        Assert.Contains(shared, shape => shape.Shape.StrokeColor == light);
        Assert.All(shared, shape => Assert.Equal(3D, shape.Shape.StrokeWidth, 3));
        Assert.Contains("stroke=\"#464646\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke=\"#B9B9B9\"", svg, StringComparison.Ordinal);
        Assert.Contains("LeftPdf", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlTables_CollapsedHiddenBorderSuppressesSharedCellEdge() {
        const string html = "<table id='hidden-conflict' style='width:100px;margin:0;table-layout:fixed;border-collapse:collapse'><tr>"
            + "<td style='border:1px solid black;border-right:5px solid red'>Left</td>"
            + "<td style='border:1px solid black;border-left:1px hidden blue'>Right</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 110D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "table#hidden-conflict:collapsed-border-v-1-0");
    }

    [Fact]
    public void HtmlTables_CollapsedBordersHonorCellRowAndTrackOriginPrecedence() {
        const string html = "<table id='origins' style='width:100px;margin:0;table-layout:fixed;border-collapse:collapse;border:2px solid purple'>"
            + "<colgroup style='border:2px solid orange'><col style='border-right:2px solid blue'></colgroup><colgroup><col></colgroup>"
            + "<tbody style='border:2px solid green'>"
            + "<tr style='border:2px solid blue'><td style='border:none;border-bottom:2px solid red'>A</td><td style='border:none'>B</td></tr>"
            + "<tr style='border:none'><td style='border:none'>C</td><td style='border:none'>D</td></tr>"
            + "</tbody></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 110D,
            ViewportHeight = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<HtmlRenderShape> shapes = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>().ToList();

        Assert.Equal(OfficeColor.Blue, Assert.Single(shapes, shape => shape.Source == "table#origins:collapsed-border-h-0-0").Shape.StrokeColor);
        Assert.Equal(OfficeColor.Red, Assert.Single(shapes, shape => shape.Source == "table#origins:collapsed-border-h-1-0").Shape.StrokeColor);
        Assert.Equal(OfficeColor.Blue, Assert.Single(shapes, shape => shape.Source == "table#origins:collapsed-border-v-1-0").Shape.StrokeColor);
        Assert.Equal(OfficeColor.Green, Assert.Single(shapes, shape => shape.Source == "table#origins:collapsed-border-h-2-1").Shape.StrokeColor);
    }

    [Fact]
    public void HtmlTables_CollapsedTableBorderPaintsOnlyResolvedOuterSegments() {
        const string html = "<table id='outer' style='width:100px;margin:0;table-layout:fixed;border-collapse:collapse;border:3px solid purple'>"
            + "<tr style='border:none;border-top:1px solid blue'><td style='border:none'>A</td><td style='border:none'>B</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 110D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<HtmlRenderShape> shapes = rendered.Pages[0].Visuals.OfType<HtmlRenderShape>().ToList();
        IReadOnlyList<HtmlRenderShape> collapsed = shapes
            .Where(shape => shape.Source != null && shape.Source.StartsWith("table#outer:collapsed-border-", StringComparison.Ordinal))
            .ToList();

        Assert.Equal(6, collapsed.Count);
        Assert.All(collapsed, shape => {
            Assert.Equal(OfficeColor.Purple, shape.Shape.StrokeColor);
            Assert.Equal(3D, shape.Shape.StrokeWidth, 3);
        });
        Assert.DoesNotContain(shapes, shape => shape.Source == "table#outer" && shape.Shape.StrokeColor == OfficeColor.Purple);
        Assert.DoesNotContain(collapsed, shape => shape.Source == "table#outer:collapsed-border-v-1-0");
    }

    [Fact]
    public void HtmlTables_InvalidCaptionSideUsesCatalogedTopFallbackAndSupportsTruth() {
        const string html = "<table id='table' style='caption-side:left;table-layout:balanced;border-collapse:merge;border-spacing:-2px;width:60px;margin:0'><caption>Caption</caption><tr><td>Cell</td></tr></table>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);
        HtmlRenderText caption = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Caption");
        HtmlRenderText cell = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Cell");

        Assert.Equal("table#table", diagnostic.Source);
        Assert.Contains("caption-side=left", diagnostic.Detail, StringComparison.Ordinal);
        Assert.Contains("table-layout=balanced", diagnostic.Detail, StringComparison.Ordinal);
        Assert.Contains("border-collapse=merge", diagnostic.Detail, StringComparison.Ordinal);
        Assert.Contains("border-spacing=-2px", diagnostic.Detail, StringComparison.Ordinal);
        Assert.True(caption.Y < cell.Y);
        Assert.Contains(HtmlRenderDiagnosticCodes.TableValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.TableValueUnsupported, out _));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(caption-side:top)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(caption-side:bottom)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(caption-side:left)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(table-layout:auto)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(table-layout:fixed)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(table-layout:balanced)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-collapse:collapse)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-spacing:2px 4px)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(border-spacing:-1px)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(border-spacing:10%)"));
    }
}
