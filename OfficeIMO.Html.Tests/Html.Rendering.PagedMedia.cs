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
