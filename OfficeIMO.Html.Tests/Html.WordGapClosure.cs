using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlWordGapClosure {
    private const string ValidPng =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";

    [Theory]
    [InlineData("ltr", 150, 300)]
    [InlineData("rtl", 300, 150)]
    public void HtmlToWord_LogicalSpacing_ResolvesAgainstEffectiveDirection(
        string direction,
        int expectedBefore,
        int expectedAfter) {
        string html = $"""
            <div dir="{direction}">
              <p style="margin-inline-start:10px;margin-inline-end:20px;padding-block:4px 6px">
                Logical spacing
              </p>
            </div>
            """;

        HtmlToWordResult conversion = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument document = conversion.Value;
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text.Contains("Logical spacing", StringComparison.Ordinal));

        Assert.Equal(expectedBefore, paragraph.IndentationBefore);
        Assert.Equal(expectedAfter, paragraph.IndentationAfter);
        Assert.Equal(60, paragraph.LineSpacingBefore);
        Assert.Equal(90, paragraph.LineSpacingAfter);
        Assert.DoesNotContain(
            conversion.Report.Diagnostics,
            diagnostic => diagnostic.Code == "UnsupportedCssDeclaration" &&
                          diagnostic.Source?.Contains("inline", StringComparison.OrdinalIgnoreCase) == true);
    }

    [Fact]
    public void HtmlToWord_LogicalSpacing_RespectsDeclarationOrderWithPhysicalProperties() {
        const string html = """
            <p style="margin-inline-start:10px;margin-left:30px">Physical last</p>
            <p style="margin-left:30px;margin-inline-start:10px">Logical last</p>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph physicalLast = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Physical last");
        WordParagraph logicalLast = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Logical last");

        Assert.Equal(450, physicalLast.IndentationBefore);
        Assert.Equal(150, logicalLast.IndentationBefore);
    }

    [Fact]
    public void HtmlToWord_LogicalPadding_RespectsDeclarationOrderWithPhysicalProperties() {
        const string html = """
            <p style="padding-left:30px;padding-inline-start:10px;padding-top:40px;padding-block-start:20px">Logical last</p>
            <p style="padding-inline-start:10px;padding-left:30px;padding-block-start:20px;padding-top:40px">Physical last</p>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph logicalLast = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Logical last");
        WordParagraph physicalLast = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Physical last");

        Assert.Equal(150, logicalLast.IndentationBefore);
        Assert.Equal(300, logicalLast.LineSpacingBefore);
        Assert.Equal(450, physicalLast.IndentationBefore);
        Assert.Equal(600, physicalLast.LineSpacingBefore);
    }

    [Fact]
    public void HtmlToWord_CssDirection_OverridesDirForLogicalSpacing() {
        const string html = """
            <style>
              .ltr { direction:ltr; }
              .ancestor { direction:ltr; }
            </style>
            <p class="ltr" dir="rtl" style="margin-inline-start:10px">CSS direction</p>
            <div class="ancestor">
              <p dir="rtl" style="margin-inline-start:10px">Own dir</p>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(document.Paragraphs, candidate => candidate.Text == "CSS direction");
        WordParagraph ownDir = Assert.Single(document.Paragraphs, candidate => candidate.Text == "Own dir");

        Assert.Equal(150, paragraph.IndentationBefore);
        Assert.Null(paragraph.IndentationAfter);
        Assert.Null(ownDir.IndentationBefore);
        Assert.Equal(150, ownDir.IndentationAfter);
    }

    [Fact]
    public void HtmlToWord_BlockContainer_PreservesOneNativeFrameAcrossParagraphs() {
        const string html = """
            <style>.frame { background-color:#abcdef; border:1px solid #123456; padding:4px 8px; }</style>
            <div class="frame"><p>First</p><p>Second</p></div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph first = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "First");
        WordParagraph second = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Second");

        Assert.Equal("ABCDEF", first.ShadingFillColorHex);
        Assert.Equal("ABCDEF", second.ShadingFillColorHex);
        Assert.Equal(BorderValues.Single, first.Borders.TopStyle);
        Assert.Null(first.Borders.BottomStyle);
        Assert.Null(second.Borders.TopStyle);
        Assert.Equal(BorderValues.Single, second.Borders.BottomStyle);
        Assert.Equal(BorderValues.Single, first.Borders.LeftStyle);
        Assert.Equal(BorderValues.Single, second.Borders.RightStyle);
        Assert.Equal("123456", first.Borders.LeftColorHex);
        Assert.Equal(120, first.IndentationBefore);
        Assert.Equal(120, second.IndentationAfter);
    }

    [Fact]
    public void HtmlToWord_BlockContainer_PreservesMoreSpecificDescendantBackground() {
        const string html = """
            <div style="background-color:#0000ff">
              <p style="background-color:#ff0000">Specific</p>
              <p>Inherited frame</p>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph specific = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Specific");
        WordParagraph inherited = Assert.Single(document.Paragraphs, paragraph => paragraph.Text == "Inherited frame");

        Assert.Equal("FF0000", specific.ShadingFillColorHex);
        Assert.Equal("0000FF", inherited.ShadingFillColorHex);
    }

    [Fact]
    public void HtmlToWord_BorderColorLonghand_DoesNotSynthesizeVisibleBorder() {
        const string html = """<p style="border-left-color:#ff0000">No border style</p>""";

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(document.Paragraphs, candidate => candidate.Text == "No border style");

        Assert.Null(paragraph.Borders.LeftStyle);
    }

    [Fact]
    public void HtmlToWord_ExactTextBackground_SurvivesSaveReloadAndHtmlExport() {
        const string html = """<p>Before <span style="background-color:#abcdef">exact</span> after</p>""";
        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph exactRun = Assert.Single(
            document.Paragraphs[0].GetRuns(),
            run => run.Text.Contains("exact", StringComparison.Ordinal));

        Assert.Equal("ABCDEF", exactRun.RunShadingFillColorHex);
        Assert.Null(exactRun.Highlight);

        using var stream = new MemoryStream();
        document.Save(stream);
        using (WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(stream.ToArray()), false)) {
            var validationErrors = new OpenXmlValidator().Validate(package).ToList();
            Assert.True(
                validationErrors.Count == 0,
                OpenXmlValidationFormatting.FormatValidationErrors(validationErrors));
        }
        using WordDocument reloaded = WordDocument.Load(new MemoryStream(stream.ToArray()));
        WordParagraph reloadedRun = Assert.Single(
            reloaded.Paragraphs[0].GetRuns(),
            run => run.Text.Contains("exact", StringComparison.Ordinal));

        Assert.Equal("ABCDEF", reloadedRun.RunShadingFillColorHex);
        Assert.Contains(
            "background-color:#abcdef",
            reloaded.ToHtml(new WordToHtmlOptions { IncludeRunHighlightStyles = true }),
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlToWord_NearestHighlightMode_ReportsColorApproximation() {
        var options = new HtmlToWordOptions {
            TextBackgroundMode = HtmlTextBackgroundMode.NearestHighlight
        };

        HtmlToWordResult conversion = HtmlConversionDocument
            .Parse("""<p><span style="background-color:#abcdef">approximate</span></p>""")
            .ToWordDocumentResult(options);
        using WordDocument document = conversion.Value;
        WordParagraph run = Assert.Single(document.Paragraphs[0].GetRuns());

        Assert.NotNull(run.Highlight);
        Assert.Equal(string.Empty, run.RunShadingFillColorHex);
        Assert.Contains(
            conversion.Report.Diagnostics,
            diagnostic => diagnostic.Code == "TextBackgroundColorApproximated");
    }

    [Fact]
    public async Task HtmlToWord_RemoteImages_UseConfiguredBoundedConcurrency() {
        int active = 0;
        int maximum = 0;
        byte[] imageBytes = Convert.FromBase64String(ValidPng);
        using var httpClient = new HttpClient(new TrackingHandler(async cancellationToken => {
            int current = Interlocked.Increment(ref active);
            UpdateMaximum(ref maximum, current);
            try {
                await Task.Delay(75, cancellationToken);
                var response = new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(imageBytes)
                };
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("image/png");
                return response;
            } finally {
                Interlocked.Decrement(ref active);
            }
        }));
        var options = new HtmlToWordOptions {
            HttpClient = httpClient,
            ImageProcessing = ImageProcessingMode.Embed,
            MaxConcurrentResourceLoads = 2
        };
        const string html = """
            <img src="https://example.test/one.png" alt="One">
            <img src="https://example.test/two.png" alt="Two">
            <img src="https://example.test/three.png" alt="Three">
            """;

        using WordDocument document = await HtmlConversionDocument.Parse(html).ToWordDocumentAsync(options);

        Assert.Equal(3, document.Images.Count);
        Assert.Equal(2, maximum);
    }

    [Fact]
    public void HtmlToWordOptions_Clone_PreservesNewConversionContracts() {
        var options = new HtmlToWordOptions {
            MaxConcurrentResourceLoads = 3,
            TextBackgroundMode = HtmlTextBackgroundMode.NearestHighlight
        };

        HtmlToWordOptions clone = options.Clone();

        Assert.Equal(3, clone.MaxConcurrentResourceLoads);
        Assert.Equal(HtmlTextBackgroundMode.NearestHighlight, clone.TextBackgroundMode);
    }

    [Fact]
    public async Task HtmlToWord_InvalidResourceConcurrency_IsRejectedBeforeNetworkAccess() {
        var options = new HtmlToWordOptions {
            ImageProcessing = ImageProcessingMode.Embed,
            MaxConcurrentResourceLoads = 0
        };

        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() =>
            HtmlConversionDocument
                .Parse("""<img src="https://example.test/image.png">""")
                .ToWordDocumentAsync(options));
    }

    [Fact]
    public async Task WordTableCell_AddHtmlAsync_PreservesTypedContainerAndArtifact() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        await cell.AddHtmlAsync(HtmlConversionDocument.Parse("""
            <h2>Cell heading</h2>
            <ul><li>First item</li><li>Second item</li></ul>
            <table><tr><td>Nested</td></tr></table>
            """));

        Assert.Contains(cell.Paragraphs, paragraph => paragraph.Text.Contains("Cell heading", StringComparison.Ordinal));
        Assert.Contains(cell.Elements, element => element is WordTable);

        using var stream = new MemoryStream();
        document.Save(stream);
        using WordDocument reloaded = WordDocument.Load(new MemoryStream(stream.ToArray()));
        WordTableCell reloadedCell = reloaded.Tables[0].Rows[0].Cells[0];

        Assert.Contains(reloadedCell.Paragraphs, paragraph => paragraph.Text.Contains("Cell heading", StringComparison.Ordinal));
        Assert.Contains(reloadedCell.Elements, element => element is WordTable);
        Assert.Contains(reloadedCell.Paragraphs, paragraph => paragraph.Text.Contains("First item", StringComparison.Ordinal));
    }

    [Fact]
    public void WordTableCell_AddHtml_AppendsTopLevelInlineSiblingsWithoutDeletingContent() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        cell.Paragraphs[0].Text = "Existing";

        cell.AddHtml(HtmlConversionDocument.Parse("""<span>First</span><span>Second</span>"""));

        Assert.Contains(cell.Paragraphs, paragraph => paragraph.Text == "Existing");
        Assert.Equal("ExistingFirstSecond", string.Concat(cell.Paragraphs.Select(paragraph => paragraph.Text)));
    }

    [Fact]
    public async Task WordTableCell_AddHtmlAsync_KeepsSectionsImagesAndSvgInCellScope() {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];
        string html = $"""
            <section><p>Scoped section</p></section>
            <img src="data:image/png;base64,{ValidPng}" alt="Scoped PNG">
            <svg xmlns="http://www.w3.org/2000/svg" width="10" height="10">
              <rect width="10" height="10" fill="#123456"/>
            </svg>
            """;

        await cell.AddHtmlAsync(HtmlConversionDocument.Parse(html));

        Assert.Contains(cell.Paragraphs, paragraph => paragraph.Text == "Scoped section");
        Assert.Equal(2, document.Images.Count);
        Assert.DoesNotContain(document.Sections[0].Paragraphs, paragraph => paragraph.Text == "Scoped section");
        Assert.DoesNotContain(document.Sections[0].Paragraphs, paragraph => paragraph.IsImage);

        using var stream = new MemoryStream();
        document.Save(stream);
        using WordDocument reloaded = WordDocument.Load(new MemoryStream(stream.ToArray()));
        WordTableCell reloadedCell = reloaded.Tables[0].Rows[0].Cells[0];

        Assert.Contains(reloadedCell.Paragraphs, paragraph => paragraph.Text == "Scoped section");
        Assert.Equal(2, reloaded.Images.Count);
    }

    [Fact]
    public async Task WordTableCell_AddHtmlAsync_UsesHeaderPartForTopLevelImage() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordHeader header = document.Sections[0].Header.Default!;
        WordTableCell cell = header.AddTable(1, 1).Rows[0].Cells[0];
        string html = $"""<img src="data:image/png;base64,{ValidPng}" alt="Header PNG">""";

        await cell.AddHtmlAsync(HtmlConversionDocument.Parse(html));

        Assert.Contains(cell.Paragraphs, paragraph => paragraph.IsImage);
        Assert.Empty(document.Images);

        using var stream = new MemoryStream();
        document.Save(stream);
        using WordDocument reloaded = WordDocument.Load(new MemoryStream(stream.ToArray()));

        WordTableCell reloadedCell = reloaded.Sections[0].Header.Default!.Tables[0].Rows[0].Cells[0];
        Assert.Contains(reloadedCell.Paragraphs, paragraph => paragraph.IsImage);
        Assert.Empty(reloaded.Images);
    }

    private static void UpdateMaximum(ref int maximum, int candidate) {
        int observed;
        while (candidate > (observed = Volatile.Read(ref maximum))) {
            if (Interlocked.CompareExchange(ref maximum, candidate, observed) == observed) {
                return;
            }
        }
    }

    private sealed class TrackingHandler : HttpMessageHandler {
        private readonly Func<CancellationToken, Task<HttpResponseMessage>> _handler;

        internal TrackingHandler(Func<CancellationToken, Task<HttpResponseMessage>> handler) {
            _handler = handler;
        }

        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken) =>
            _handler(cancellationToken);
    }
}
