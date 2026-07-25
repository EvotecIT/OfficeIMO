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
