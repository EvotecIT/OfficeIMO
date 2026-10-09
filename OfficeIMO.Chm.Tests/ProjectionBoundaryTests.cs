using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Providers;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Chm;
using OfficeIMO.Reader.Html;
using System.Threading;

namespace OfficeIMO.Chm.Tests;

public sealed class ProjectionBoundaryTests {
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
