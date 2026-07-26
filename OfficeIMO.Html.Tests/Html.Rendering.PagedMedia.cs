using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_Paged_ResolvesRunningStringsFromPageLocalAssignments() {
        const string html = """
            <style>
              @page {
                size: 3in 2in;
                margin: 24px;
                @top-left { content: "start=" string(chapter, start); }
                @top-center { content: "first=" string(chapter); }
                @top-right { content: "last=" string(chapter, last); }
                @bottom-center { content: "except=" string(chapter, first-except); }
              }
              h1, h2 { string-set: chapter content(); margin:0; font-size:12px; line-height:14px; }
            </style>
            <h1>Opening Chapter</h1>
            <p style="margin:0">Opening body</p>
            <h2 style="break-before:page">Second Chapter</h2>
            <p style="margin:0">Second body</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(4D, 4D),
            Margins = HtmlRenderMargins.All(10D)
        });

        Assert.Equal(2, rendered.Pages.Count);
        IReadOnlyList<string> first = rendered.Pages[0].Visuals.OfType<HtmlRenderText>()
            .Where(text => text.SemanticRole == "page-margin")
            .Select(text => text.Text)
            .ToList();
        IReadOnlyList<string> second = rendered.Pages[1].Visuals.OfType<HtmlRenderText>()
            .Where(text => text.SemanticRole == "page-margin")
            .Select(text => text.Text)
            .ToList();

        Assert.Contains("start=", first);
        Assert.Contains("first=Opening Chapter", first);
        Assert.Contains("last=Opening Chapter", first);
        Assert.Contains("except=", first);
        Assert.Contains("start=Opening Chapter", second);
        Assert.Contains("first=Second Chapter", second);
        Assert.Contains("last=Second Chapter", second);
        Assert.Contains("except=", second);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.PageMarginContentUnsupported);
    }

    [Fact]
    public void HtmlRender_Paged_RunningStringsSupportLiteralAndAttributeContent() {
        const string html = """
            <style>
              @page {
                size: 3in 2in;
                margin: 24px;
                @top-center { content: string(section); }
              }
              h2 { string-set: section "Part " attr(data-part) ": " content(); margin:0; }
            </style>
            <h2 data-part="IV" style='string-set:section "Part " attr(data-part) ": " content()'>Maintenance</h2><p>Body</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged
        });

        HtmlRenderText margin = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.SemanticRole == "page-margin");
        Assert.Equal("Part IV: Maintenance", margin.Text);
    }

    [Theory]
    [InlineData("\\41", "A")]
    [InlineData("\\000041 ", "A")]
    [InlineData("\\1F600", "\ud83d\ude00")]
    public void HtmlRender_Paged_RunningStringLiteralsDecodeCssHexadecimalEscapes(
        string literal,
        string expected) {
        string html = """
            <style>
              @page { size:3in 2in; margin:24px; @top-center { content:string(section); } }
            </style>
            """
            + "<h2 style=\"string-set:section '" + literal + "'\">Body</h2>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        HtmlRenderText margin = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.SemanticRole == "page-margin");
        Assert.Equal(expected, margin.Text);
    }

    [Fact]
    public void HtmlRender_Paged_DecodesEscapedRunningStringIdentifiersOnAssignmentAndLookup() {
        const string html = """
            <style>
              @page { size:3in 2in; margin:24px; @top-center { content:string(sec\74 ion); } }
            </style>
            <h2 style="string-set:sec\74 ion 'Escaped identifier'">Body</h2>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        HtmlRenderText margin = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.SemanticRole == "page-margin");
        Assert.Equal("Escaped identifier", margin.Text);
    }

    [Fact]
    public void HtmlRender_Paged_PreservesRunningStringFlowOrderAcrossColumns() {
        const string html = """
            <style>
              @page {
                size:4in 3in;
                margin:24px;
                @top-left { content:"first=" string(section, first); }
                @top-right { content:"last=" string(section, last); }
              }
              h2 { margin:0; font-size:12px; line-height:14px; }
            </style>
            <div style="column-count:2;column-fill:auto;height:60px;width:220px">
              <div style="height:40px"></div>
              <h2 style="string-set:section 'First column'">First</h2>
              <div style="height:20px"></div>
              <h2 style="string-set:section 'Second column'">Second</h2>
            </div>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        IReadOnlyList<string> margins = rendered.Pages[0].Visuals
            .OfType<HtmlRenderText>()
            .Where(text => text.SemanticRole == "page-margin")
            .Select(text => text.Text)
            .ToList();
        Assert.Contains("first=First column", margins);
        Assert.Contains("last=Second column", margins);
    }

    [Theory]
    [InlineData("<div style='display:flex'><h2 style='string-set:section content()'>Flex section</h2></div>", "Flex section")]
    [InlineData("<div style='display:grid;grid-template-columns:1fr'><h2 style='string-set:section content()'>Grid section</h2></div>", "Grid section")]
    [InlineData("<div style='column-count:2'><h2 style='string-set:section content()'>Column section</h2><p>Body</p></div>", "Column section")]
    [InlineData("<table><tr><td><h2 style='string-set:section content()'>Table section</h2></td></tr></table>", "Table section")]
    public void HtmlRender_Paged_PropagatesRunningStringsThroughSpecializedContainers(
        string body,
        string expected) {
        string html = "<style>@page{size:3in 2in;margin:24px;@top-center{content:string(section)}}h2{margin:0;font-size:12px;line-height:14px}</style>" + body;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Contains(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.SemanticRole == "page-margin" && text.Text == expected);
    }

    [Theory]
    [InlineData("<h2 style='position:absolute;top:0;left:0;string-set:section content()'>Positioned section</h2>")]
    [InlineData("<h2 style='position:fixed;top:0;left:0;string-set:section content()'>Positioned section</h2>")]
    [InlineData("<div style='position:relative;height:40px'><h2 style='position:absolute;top:0;left:0;string-set:section content()'>Positioned section</h2></div>")]
    public void HtmlRender_Paged_OrdersPositionedRunningStringsByResolvedPageOffset(string positionedBody) {
        string html = """
            <style>
              @page { size:3in 2in; margin:24px; @top-center { content:string(section); } }
              h2 { margin:0; font-size:12px; line-height:14px; }
            </style>
            """
            + positionedBody
            + "<div style='height:48px'></div><h2 style='string-set:section content()'>Flow section</h2>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Contains(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.SemanticRole == "page-margin" && text.Text == "Positioned section");
    }

    [Fact]
    public void HtmlRender_Paged_DoesNotRecordPositionedRunningStringAtStaticInlineAnchor() {
        const string html = """
            <style>
              @page { size:3in 3in; margin:24px; @top-center { content:string(section, first); } }
              p, h2 { margin:0; font-size:12px; line-height:14px; }
            </style>
            <p>Anchor <span style="position:absolute;top:100px;left:0;string-set:section 'Positioned'">Positioned</span></p>
            <div style="height:48px"></div>
            <h2 style="string-set:section 'Flow'">Flow</h2>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Contains(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.SemanticRole == "page-margin" && text.Text == "Flow");
    }

    [Theory]
    [InlineData("string-set:section 'old';string-set:section 'new'", "new")]
    [InlineData("string-set:section 'old' !important;string-set:section 'new'", "old")]
    [InlineData("string-set:section 'old';string-set:section 'new' !important", "new")]
    [InlineData("string-set:section 'old' !important;string-set:section 'new' !important", "new")]
    public void HtmlRender_Paged_RetainedStringSetDeclarationsRespectSourceOrderAndImportance(
        string declarations,
        string expected) {
        string html = "<style>@page{size:3in 2in;margin:24px;@top-center{content:string(section)}}"
            + ".target{" + declarations + "}</style><p class='target'>Body</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Contains(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.SemanticRole == "page-margin" && text.Text == expected);
    }

    [Fact]
    public void HtmlRender_Paged_CollectsRunningStringsFromInlineElements() {
        const string html = """
            <style>
              @page { size:3in 2in; margin:24px; @top-center { content:string(section); } }
              p { margin:0; font-size:12px; line-height:14px; }
            </style>
            <p>Before <span style="string-set:section content()">Inline section</span> after</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Contains(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.SemanticRole == "page-margin" && text.Text == "Inline section");
    }

    [Theory]
    [InlineData("A <span style=\"string-set:section content()\"></span> B", "A B")]
    [InlineData("<span style=\"string-set:section 'Start'\"></span> A", "A")]
    [InlineData("A <span style=\"string-set:section 'End'\"></span>", "A")]
    public void HtmlRender_InlineRunningStringMarkersRemainTransparentToWhitespaceCollapsing(
        string body,
        string expected) {
        string html = "<p style=\"margin:0;font-size:12px;line-height:14px\">" + body + "</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        string text = string.Concat(rendered.Pages[0].Visuals
            .OfType<HtmlRenderText>()
            .Where(visual => visual.SemanticRole != "page-margin")
            .Select(visual => visual.Text));
        Assert.Equal(expected, text);
    }

    [Fact]
    public void HtmlRender_InlineRunningStringMarkersDoNotCreateLineOccupancy() {
        const string html = """
            <style>
            @page { size:3in 2in; margin:24px; @top-center { content:string(section) } }
            div { margin:0; font-size:12px; line-height:14px }
            </style>
            <div>Before</div>
            <div><span style="string-set:section 'Marker only'"></span></div>
            <div>After</div>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        HtmlRenderText before = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.Text == "Before");
        HtmlRenderText after = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.Text == "After");
        Assert.InRange(after.Y - before.Y, 14D, 14.02D);
        Assert.Contains(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.SemanticRole == "page-margin" && text.Text == "Marker only");
    }

    [Theory]
    [InlineData("<div style='float:left'><span style='string-set:section content()'>Float section</span></div><p>Body</p>", "Float section")]
    [InlineData("<p><span style='display:inline-block'><span style='string-set:section content()'>Inline block section</span></span></p>", "Inline block section")]
    [InlineData("<p><span style='display:inline-flex'><span style='string-set:section content()'>Inline flex section</span></span></p>", "Inline flex section")]
    [InlineData("<p><span style='display:inline-grid'><span style='string-set:section content()'>Inline grid section</span></span></p>", "Inline grid section")]
    public void HtmlRender_Paged_PropagatesRunningStringsThroughFloatsAndInlineAtomicBoxes(
        string body,
        string expected) {
        string html = "<style>@page{size:3in 2in;margin:24px;@top-center{content:string(section)}}"
            + "p,div,span{margin:0;font-size:12px;line-height:14px}</style>"
            + body;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Contains(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.SemanticRole == "page-margin" && text.Text == expected);
    }

    [Fact]
    public void HtmlRender_Paged_OmitsRunningStringsBeyondTheConfiguredCharacterLimit() {
        const string html = """
            <style>
              @page { @top-center { content: string(section); } }
              h2 { string-set: section content(); }
            </style>
            <h2>Unbounded descendant text</h2>
            """;
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            MaxRunningStringCharacters = 8
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.RunningStringLimitExceeded);
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text =>
            text.SemanticRole == "page-margin" && text.Text.Contains("Unbounded", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlRender_Paged_ChargesLiteralRunningStringsToTheOperationWideBudget() {
        string repeated = string.Concat(Enumerable.Repeat(
            "<h2 style='string-set:section \"abcdefghijklmnopqrstuvwxyz0123456789\"'>Heading</h2>",
            20));
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            MaxLayoutOperations = 200,
            MaxRunningStringCharacters = 128
        };

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(repeated), options));

        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), exception.LimitSource);
    }
}
