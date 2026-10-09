using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Providers;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Chm;
using OfficeIMO.Reader.Html;
using System.Threading;

namespace OfficeIMO.Chm.Tests;

public sealed class ProjectionBoundaryTests {
    [Theory]
    [InlineData("src", false)]
    [InlineData("data-src", false)]
    [InlineData("srcset", false)]
    [InlineData("data-srcset", false)]
    [InlineData("src", true)]
    [InlineData("data-src", true)]
    [InlineData("srcset", true)]
    [InlineData("data-srcset", true)]
    public void ReaderImageAssetsHonorTheResourcePolicy(string attribute, bool chm) {
        string pixel = "data:image/png;base64," + System.Convert.ToBase64String(ChmDocument.Load(ChmFixture.Book()).FindEntry("/pixel.png")!.GetBytes());
        string value = pixel + (attribute.EndsWith("srcset", StringComparison.Ordinal) ? " 1x" : "");
        string html = "<p>Text</p><img " + attribute + "='" + value + "' alt='Pixel'>";
        var options = ReaderHtmlOptions.CreateOfficeIMOProfile();
        options.HtmlToMarkdownOptions!.ResourceUrlPolicy = HtmlUrlPolicy.CreateWebResourceProfile();
        var rejected = Read(html, options, chm);
        Assert.Empty(rejected.Assets);
        Assert.DoesNotContain(rejected.Visuals, visual => visual.SourceName?.StartsWith("data:", StringComparison.Ordinal) == true);
        options.HtmlToMarkdownOptions.ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile();
        Assert.Single(Read(html, options, chm).Assets);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ReaderMediaReferencesHonorTheResourcePolicy(bool childSource, bool chm) {
        const string uri = "https://example.invalid/clip.mp4";
        string html = childSource ? "<video><source src='" + uri + "'></video>" : "<video src='" + uri + "'></video>";
        var options = ReaderHtmlOptions.CreateOfficeIMOProfile();
        options.HtmlToMarkdownOptions!.ResourceUrlPolicy = HtmlUrlPolicy.CreateWebResourceProfile();
        options.HtmlToMarkdownOptions.ResourceUrlPolicy.AllowedUrlSchemes.Clear();
        var rejected = Read(html, options, chm);
        Assert.DoesNotContain(rejected.Visuals, visual => visual.SourceName == uri || visual.Content.Contains(uri));
        options.HtmlToMarkdownOptions.ResourceUrlPolicy.AllowedUrlSchemes.Add("https");
        Assert.Contains(Read(html, options, chm).Visuals, visual => visual.SourceName == uri);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" id='forged'")]
    public void AuthoredTopicAnnotationsDoNotOwnEpubNavigation(string attributes) {
        var book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/first.html"] = ChmFixture.Html("<section data-chm-topic='/first.html'" + attributes + "><p>Text</p></section>")
        }));
        var imported = book.ToEpubPublicationResult();
        Assert.True(imported.Succeeded);
        using var stream = new MemoryStream(imported.Publication.Write().Value);
        Assert.Equal("chm-topic-1", Assert.Single(OfficeIMO.Epub.EpubDocument.Load(stream).TableOfContents).Fragment);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BodyStylesReportAnchorAndCombinedCascadeChanges(bool multiple) {
        var files = new Dictionary<string, byte[]> {
            ["/first.html"] = ChmFixture.Html("<body><style>#note { color: red; }</style><p id='note'>Text</p></body>")
        };
        if (multiple) files["/second.html"] = ChmFixture.Html("<p>Second</p>");
        var projected = ChmDocument.Load(ChmFixture.Archive(files)).ToHtmlDocumentResult();
        Assert.Contains("#note", projected.Value.SourceHtml);
        Assert.NotNull(projected.Value.Document.QuerySelector("#chm-topic-1-note"));
        Assert.True(projected.Report.HasLoss);
        Assert.Contains(projected.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "CHM_STYLE_IDENTIFIERS");
        Assert.Equal(multiple, projected.Report.FidelityDiagnostics.Any(diagnostic => diagnostic.Code == "CHM_COMBINED_STYLE_SCOPE"));
    }

    [Fact]
    public void HeadingDemotionPreservesInterleavedSiblingAndNestedContentOrder() {
        string html = string.Concat(Enumerable.Range(0, 100).Select(index =>
            "<p>Before " + index + "</p><h2 id='h" + index + "'><b>Heading " + index + "</b></h2>"));
        var book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> { ["/first.html"] = ChmFixture.Html(html) }));
        var section = book.ToHtmlDocumentResult().Value.Document.QuerySelector("section[data-chm-topic]")!;
        var children = section.Children.Skip(1).ToArray();
        Assert.Equal(200, children.Length);
        for (int index = 0; index < 100; index++) {
            Assert.Equal("Before " + index, children[index * 2].TextContent);
            Assert.Equal("h3", children[index * 2 + 1].LocalName);
            Assert.Equal("Heading " + index, children[index * 2 + 1].QuerySelector("b")!.TextContent);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CanonicalHtmlReaderPassesCancellationToTheRegisteredParser(bool rich) {
        using var cancelled = new CancellationTokenSource();
        var parser = new CancellingParser(cancelled);
        var reader = new OfficeDocumentReaderBuilder().AddHtmlHandler(new ReaderHtmlOptions {
            ConversionOptions = new HtmlConversionDocumentOptions { ParserProvider = parser }
        }).Build();
        using var stream = new MemoryStream(ChmFixture.Html("<p>Text</p>"));
        Assert.ThrowsAny<OperationCanceledException>(() => {
            if (rich) reader.ReadDocument(stream, "topic.html", cancellationToken: cancelled.Token);
            else reader.Read(stream, "topic.html", cancellationToken: cancelled.Token).ToArray();
        });
        Assert.Equal(cancelled.Token, parser.ReceivedToken);
    }

    [Fact]
    public void OwnedHtmlSnapshotHonorsCancellationWithoutMutatingTheCallerTree() {
        var source = ChmDocument.Load(TwoTopics()).ToHtmlDocumentResult().Value.Document.CloneAttached();
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        OperationCanceledException error = Assert.ThrowsAny<OperationCanceledException>(() =>
            HtmlConversionDocument.FromDocument(source, null, cancelled.Token));
        Assert.Equal(cancelled.Token, error.CancellationToken);
        Assert.False(source.IsReadOnly);
        var captured = HtmlConversionDocument.FromDocument(source, null, CancellationToken.None);
        source.QuerySelector("p")!.TextContent = "Changed";
        Assert.Equal("First", captured.Document.QuerySelector("p")!.TextContent);
    }

    [Theory]
    [InlineData("reader")]
    [InlineData("pdf")]
    public void AggregateNodeBudgetIncludesEverySelectedTopic(string route) {
        byte[] archive = TwoTopics();
        var options = new ChmConversionOptions { MaxHtmlNodes = 9 };
        Assert.Throws<HtmlDomLimitException>(() => Convert(archive, route, options));
        options.TopicPaths = new[] { "/first.html" };
        Convert(archive, route, options);
        options.TopicPaths = null;
        options.MaxHtmlNodes = 10;
        Convert(archive, route, options);
    }

    [Theory]
    [InlineData("reader")]
    [InlineData("pdf")]
    public void AggregateNodeBudgetCountsTextAndCommentNodes(string route) {
        byte[] archive = ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/first.html"] = ChmFixture.Html("<p>Text<!-- inert comment --></p>")
        });
        Assert.Throws<HtmlDomLimitException>(() => Convert(archive, route, new ChmConversionOptions { MaxHtmlNodes = 5 }));
        Convert(archive, route, new ChmConversionOptions { MaxHtmlNodes = 6 });
    }

    [Fact]
    public void HeadingDemotionReportsRetainedCssSelectorChanges() {
        ChmDocument book = ChmDocument.Load(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/first.html"] = ChmFixture.Html("<style>h1 { color: red; } .article > h2 { color: blue; }</style>" +
                "<div class='article'><h1 id='title'>Title</h1><h2>Detail</h2></div>")
        }));
        var html = book.ToHtmlDocumentResult();
        Assert.Equal("h2", html.Value.Document.QuerySelector("#chm-topic-1-title")!.LocalName);
        Assert.Contains("h1 { color: red; }", html.Value.SourceHtml);
        Assert.Contains(html.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "CHM_HEADING_LEVELS"
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation && diagnostic.Location == "/first.html");
        Assert.Contains(book.ToMarkdownResult().Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "CHM_HEADING_LEVELS");
        Assert.Contains(book.ToEpubPublicationResult().Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "CHM_HEADING_LEVELS");
    }

    private static byte[] TwoTopics() => ChmFixture.Archive(new Dictionary<string, byte[]> {
        ["/first.html"] = ChmFixture.Html("<p>First</p>"),
        ["/second.html"] = ChmFixture.Html("<p>Second</p>")
    });

    private static OfficeDocumentReadResult Read(string html, ReaderHtmlOptions options, bool chm) {
        var builder = new OfficeDocumentReaderBuilder();
        if (chm) builder.AddChmHandler(new ReaderChmOptions { HtmlOptions = options });
        else builder.AddHtmlHandler(options);
        byte[] bytes = chm ? ChmFixture.Archive(new Dictionary<string, byte[]> { ["/first.html"] = ChmFixture.Html(html) }) : ChmFixture.Html(html);
        using var stream = new MemoryStream(bytes);
        return builder.Build().ReadDocument(stream, chm ? "manual.chm" : "topic.html");
    }

    private static void Convert(byte[] archive, string route, ChmConversionOptions options) {
        if (route == "pdf") {
            ChmDocument.Load(archive).ToPdfBytesResult(options);
        } else {
            var reader = new OfficeDocumentReaderBuilder().AddChmHandler(new ReaderChmOptions { ConversionOptions = options }).Build();
            using var stream = new MemoryStream(archive);
            reader.ReadDocument(stream, "manual.chm");
        }
    }

    private sealed class CancellingParser : IHtmlParserProvider {
        private readonly CancellationTokenSource _source;
        internal CancellingParser(CancellationTokenSource source) { _source = source; }
        public string Id => "cancelling-reader-test";
        internal CancellationToken ReceivedToken { get; private set; }
        public HtmlDocument ParseDocument(string source, HtmlParseOptions options, CancellationToken cancellationToken = default) {
            ReceivedToken = cancellationToken;
            HtmlDocument parsed = AngleSharpHtmlParser.Instance.ParseDocument(source, options, cancellationToken);
            _source.Cancel();
            return parsed;
        }
        public HtmlDocumentFragment ParseFragment(string source, HtmlElement contextElement, HtmlParseOptions options, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();
    }
}
