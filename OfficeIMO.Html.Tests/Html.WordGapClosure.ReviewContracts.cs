using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using M = DocumentFormat.OpenXml.Math;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Fact]
    public void HtmlToWord_CssWideFrameValues_ResetEarlierDeclarations() {
        const string html = """
            <div style="background-color:red;background-color:initial">Background reset</div>
            <p style="border:1px solid red;border:unset">Border reset</p>
            <p style="border-left:1px solid red;border-left:revert-layer">Side reset</p>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph background = Assert.Single(
            document.Paragraphs,
            paragraph => paragraph.Text == "Background reset");
        WordParagraph border = Assert.Single(
            document.Paragraphs,
            paragraph => paragraph.Text == "Border reset");
        WordParagraph side = Assert.Single(
            document.Paragraphs,
            paragraph => paragraph.Text == "Side reset");

        Assert.True(string.IsNullOrEmpty(background.ShadingFillColorHex));
        Assert.Null(border.Borders.TopStyle);
        Assert.Null(border.Borders.LeftStyle);
        Assert.Null(side.Borders.LeftStyle);
    }

    [Fact]
    public void HtmlToWord_FrameValues_CanExplicitlyInheritFromTheParent() {
        const string html = """
            <div style="background-color:#abcdef;border:1px solid #123456">
              <p style="background-color:inherit;border:inherit">Inherited frame</p>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Inherited frame");

        Assert.Equal("ABCDEF", paragraph.ShadingFillColorHex);
        Assert.Equal(BorderValues.Single, paragraph.Borders.TopStyle);
        Assert.Equal("123456", paragraph.Borders.LeftColorHex);
    }

    [Fact]
    public void HtmlToWord_NegativeLogicalMargins_AreAppliedToParagraphIndentation() {
        const string html = """
            <p style="margin-inline-start:-10px">LTR start</p>
            <p dir="rtl" style="margin-inline-start:-12px">RTL start</p>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph ltr = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "LTR start");
        WordParagraph rtl = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "RTL start");

        Assert.Equal(-150, ltr.IndentationBefore);
        Assert.Null(ltr.IndentationAfter);
        Assert.Null(rtl.IndentationBefore);
        Assert.Equal(-180, rtl.IndentationAfter);
    }

    [Theory]
    [InlineData("div")]
    [InlineData("address")]
    [InlineData("dl")]
    public void HtmlToWord_MixedContainerTextFollowsItsLastBlockChild(string tagName) {
        string html = $"<{tagName}><p>One</p><p>Two</p>tail</{tagName}>";

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();

        Assert.Equal(
            new[] { "One", "Two", "tail" },
            document.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
    }

    [Fact]
    public void HtmlToWord_EmptyBlock_UsesTheWinningPageBreakDeclaration() {
        const string html = """
            <div style="break-before:page;break-before:auto"></div>
            <div style="break-before:page!important;break-before:auto"></div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(document.Paragraphs);

        Assert.True(paragraph.PageBreakBefore);
    }

    [Theory]
    [InlineData("section")]
    [InlineData("article")]
    [InlineData("aside")]
    [InlineData("nav")]
    [InlineData("header")]
    [InlineData("footer")]
    [InlineData("main")]
    public void HtmlToWord_StyledEmptySemanticBlock_MaterializesItsWordFrame(string tagName) {
        string html = $"<{tagName} style=\"border:1px solid #123456;background-color:#abcdef;padding:4px\"></{tagName}>";

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(document.Paragraphs);

        Assert.Equal("ABCDEF", paragraph.ShadingFillColorHex);
        Assert.Equal(BorderValues.Single, paragraph.Borders.TopStyle);
        Assert.Equal("123456", paragraph.Borders.TopColorHex);
        Assert.Equal(60, paragraph.IndentationBefore);
        Assert.Equal(60, paragraph.IndentationAfter);
    }

    [Fact]
    public void HtmlToWord_ZeroWidthBlockBorders_RemainInvisible() {
        const string html = """
            <p style="border:0px solid red">Zero shorthand</p>
            <p style="border-left:1px solid red;border-left-width:0px">Zero longhand</p>
            <p style="border:0 solid red">Unitless zero</p>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph shorthand = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Zero shorthand");
        WordParagraph longhand = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Zero longhand");
        WordParagraph unitless = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Unitless zero");

        Assert.Null(shorthand.Borders.TopStyle);
        Assert.Null(shorthand.Borders.LeftStyle);
        Assert.Null(longhand.Borders.LeftStyle);
        Assert.Null(unitless.Borders.TopStyle);
    }

    [Fact]
    public void WordTableCell_AddHtml_BelowTableCaptionReusesTheTrailingParagraph() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        var options = new HtmlToWordOptions { TableCaptionPosition = TableCaptionPosition.Below };

        cell.AddHtml(
            HtmlConversionDocument.Parse(
                """<table><caption>Nested caption</caption><tr><td>Value</td></tr></table>"""),
            options);

        OpenXmlElement[] elements = cell._tableCell.ChildElements
            .Where(element => element is Table or Paragraph)
            .ToArray();
        Assert.Collection(
            elements,
            element => Assert.IsType<Table>(element),
            element => Assert.Equal("Nested caption", Assert.IsType<Paragraph>(element).InnerText));
    }

    [Fact]
    public void WordTableCell_AddHtml_RejectedImageDoesNotLeaveAnEmptyParagraph() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.Paragraphs[0].Text = "Existing";

        cell.AddHtml(HtmlConversionDocument.Parse("""<img alt="">"""));

        Paragraph paragraph = Assert.Single(cell._tableCell.Elements<Paragraph>());
        Assert.Equal("Existing", paragraph.InnerText);
    }

    [Fact]
    public void WordTableCell_AddHtml_PreservesFormattedEmptyParagraphs() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.Paragraphs[0].PageBreakBefore = true;

        cell.AddHtml(HtmlConversionDocument.Parse("""<p>Added</p>"""));

        Paragraph[] paragraphs = cell._tableCell.Elements<Paragraph>().ToArray();
        Assert.Equal(2, paragraphs.Length);
        Assert.NotNull(paragraphs[0].ParagraphProperties?.GetFirstChild<PageBreakBefore>());
        Assert.Equal(string.Empty, paragraphs[0].InnerText);
        Assert.Equal("Added", paragraphs[1].InnerText);
    }

    [Fact]
    public void HtmlToWord_InheritedLogicalPadding_UsesTheParentComputedValue() {
        const string html = """
            <div style="padding-inline-start:10px">
              <p style="padding-inline-start:inherit">Inherited padding</p>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Inherited padding");

        Assert.Equal(300, paragraph.IndentationBefore);
        Assert.Null(paragraph.IndentationAfter);
    }

    [Fact]
    public void HtmlToWord_NegativeLogicalPadding_DoesNotOverrideValidPhysicalPadding() {
        const string html = """
            <p style="padding-left:20px;padding-inline-start:-5px">Negative longhand</p>
            <p style="padding-top:20px;padding-block:-5px 4px">Negative pair</p>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph longhand = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Negative longhand");
        WordParagraph pair = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Negative pair");
        CssStyleMapper.CssProperties negativeMargin = CssStyleMapper.ParseStyles("margin:-5px");

        Assert.Equal(300, longhand.IndentationBefore);
        Assert.Equal(300, pair.LineSpacingBefore);
        Assert.Null(pair.LineSpacingAfter);
        Assert.Equal(-75, negativeMargin.MarginLeft);
    }

    [Fact]
    public void HtmlToWord_StyledEmptyBlock_MaterializesItsWordFrame() {
        const string html = """
            <div style="border:1px solid #123456;background-color:#abcdef;padding:4px"></div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(document.Paragraphs);

        Assert.Equal("ABCDEF", paragraph.ShadingFillColorHex);
        Assert.Equal(BorderValues.Single, paragraph.Borders.TopStyle);
        Assert.Equal("123456", paragraph.Borders.TopColorHex);
        Assert.Equal(60, paragraph.IndentationBefore);
        Assert.Equal(60, paragraph.IndentationAfter);
    }

    [Fact]
    public void WordToHtml_VisibleHighlight_TakesPrecedenceOverExactRunShading() {
        using WordDocument document = WordDocument.Create();
        WordParagraph run = document.AddParagraph("Layered");
        run.RunShadingFillColorHex = "ABCDEF";
        run.Highlight = HighlightColorValues.Cyan;

        string html = document.ToHtml(new WordToHtmlOptions { IncludeRunHighlightStyles = true });

        Assert.Contains("background-color:#00ffff", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("background-color:#abcdef", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlToWord_CssWideBoxValues_ResetEarlierPhysicalAndLogicalSpacing() {
        const string html = """
            <p style="margin-left:10px;margin-left:initial;padding-right:12px;padding-right:unset">Physical reset</p>
            <p style="margin-inline-start:10px;margin-inline-start:revert;padding-block:5px;padding-block:revert-layer">Logical reset</p>
            <p style="margin:10px;margin:inherit;padding:8px;padding:initial">Shorthand reset</p>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph physical = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Physical reset");
        WordParagraph logical = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Logical reset");
        WordParagraph shorthand = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Shorthand reset");

        Assert.Null(physical.IndentationBefore);
        Assert.Null(physical.IndentationAfter);
        Assert.Null(logical.IndentationBefore);
        Assert.Null(logical.LineSpacingBefore);
        Assert.Null(logical.LineSpacingAfter);
        Assert.Null(shorthand.IndentationBefore);
        Assert.Null(shorthand.IndentationAfter);
        Assert.Null(shorthand.LineSpacingBefore);
        Assert.Null(shorthand.LineSpacingAfter);
    }

    [Fact]
    public void HtmlToWord_BlockBorders_PreserveResetComponentsAndUseCurrentTextColor() {
        const string html = """
            <p style="border-left:1px red;border-left-style:solid">Restored components</p>
            <p style="color:#123456;border:1px solid">Current color</p>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph restored = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Restored components");
        WordParagraph currentColor = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Current color");

        Assert.Equal(BorderValues.Single, restored.Borders.LeftStyle);
        Assert.Equal((uint)6, restored.Borders.LeftSize?.Value);
        Assert.Equal("FF0000", restored.Borders.LeftColorHex);
        Assert.Equal(BorderValues.Single, currentColor.Borders.TopStyle);
        Assert.Equal("123456", currentColor.Borders.TopColorHex);
    }

    [Fact]
    public void WordToHtml_ExactTextBackground_ExportsBesideAnEquation() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        paragraph._paragraph.Append(new Hyperlink(
            new Run(
                new RunProperties(new Shading {
                    Val = ShadingPatternValues.Clear,
                    Fill = "ABCDEF",
                }),
                new Text("shaded")),
            new M.OfficeMath(new M.Run(new M.Text("x")))) {
            Anchor = "target",
        });

        string html = document.ToHtml(new WordToHtmlOptions { IncludeRunHighlightStyles = true });

        Assert.Contains("background-color:#abcdef", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<math", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task HtmlToWord_RemoteImages_DeduplicateCanonicalUris() {
        int requestCount = 0;
        byte[] imageBytes = Convert.FromBase64String(ValidPng);
        using var httpClient = new HttpClient(new TrackingHandler(_ => {
            Interlocked.Increment(ref requestCount);
            var response = new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new ByteArrayContent(imageBytes),
            };
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("image/png");
            return Task.FromResult(response);
        }));
        var options = new HtmlToWordOptions {
            HttpClient = httpClient,
            ImageProcessing = ImageProcessingMode.Embed,
            MaxConcurrentResourceLoads = 2,
        };
        const string html = """
            <img src="https://EXAMPLE.test:443/shared.png" alt="First">
            <img src="https://example.test/shared.png" alt="Second">
            """;

        using WordDocument document = await HtmlConversionDocument.Parse(html).ToWordDocumentAsync(options);

        Assert.Equal(1, Volatile.Read(ref requestCount));
        Assert.NotEmpty(document.Images);
    }

    [Fact]
    public void WordTableCell_AddHtml_ListReplacesTheFreshPlaceholder() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        cell.AddHtml(HtmlConversionDocument.Parse("""<ul><li>Only item</li></ul>"""));

        Paragraph paragraph = Assert.Single(cell._tableCell.Elements<Paragraph>());
        Assert.Equal("Only item", paragraph.InnerText);
        Assert.Contains(cell.Paragraphs, item => item.IsListItem && item.Text == "Only item");
    }
}
