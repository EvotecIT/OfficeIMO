using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Providers;
using OfficeIMO.Markdown.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlFoundationProviderContractTests {
    [Theory]
    [InlineData("\u00a0")]
    [InlineData("\u2003")]
    [InlineData("\u202f")]
    public void ParserRecoveredUnicodeAttributeNamesSurviveOwnedEditingAndCallbackSnapshots(string name) {
        string source = "<p " + name + "='kept'>Text</p>";
        HtmlDocument parsed = AngleSharpHtmlParser.Instance.Parse(source, new HtmlParseOptions());
        Assert.Equal("kept", parsed.QuerySelector("p")!.GetAttribute(name));
        HtmlConversionDocument conversion = HtmlConversionDocument.Parse(source);
        Assert.Equal("kept", conversion.Document.QuerySelector("p")!.GetAttribute(name));
        HtmlConversionDocument edited = conversion.Edit(document => document.QuerySelector("p")!.SetAttribute(name, "edited"));
        Assert.Equal("edited", edited.Document.QuerySelector("p")!.GetAttribute(name));
        Assert.Equal("kept", conversion.Document.QuerySelector("p")!.GetAttribute(name));
        Assert.Contains("Text", edited.ToMarkdown());
        var native = NativeDomBridge.GetNativeDocument(edited.Document);
        try { Assert.Equal("edited", NativeDomBridge.Wrap(native.QuerySelector("p")!).GetAttribute(name)); }
        finally { NativeDomBridge.ReleaseCallbackSnapshot(native); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RootlessDocumentsNormalizeAndConvertWithCommentPolicy(bool includeNonElementNodes) {
        var document = new HtmlDocument(AngleSharpDomServices.Instance, "authored");
        if (includeNonElementNodes) {
            document.AppendChild(document.CreateDocumentType("html"));
            document.AppendChild(document.CreateComment("retained"));
        }
        foreach (bool bodyOnly in new[] { false, true }) {
            Assert.Equal(string.Empty, HtmlNormalizer.Normalize(document, new HtmlNormalizationOptions { UseBodyContentsOnly = bodyOnly, PreserveComments = false }));
            Assert.Equal(includeNonElementNodes ? "<!--retained-->" : string.Empty,
                HtmlNormalizer.Normalize(document, new HtmlNormalizationOptions { UseBodyContentsOnly = bodyOnly, PreserveComments = true }));
        }
        HtmlConversionDocument conversion = HtmlConversionDocument.FromDocument(document);
        Assert.Null(conversion.Document.DocumentElement);
        Assert.Empty(conversion.ToMarkdown().Trim());
        Assert.NotNull(conversion.SemanticDocument);
        Assert.Empty(conversion.CreateDocumentForConversion().Body!.TextContent);
    }

    [Fact]
    public void RemovingTheDocumentElementProducesAnEmptyConversionWithoutChangingTheSource() {
        HtmlConversionDocument original = HtmlConversionDocument.Parse("<!doctype html><!--retained--><p>Before</p>");
        HtmlConversionDocument edited = original.Edit(document => document.DocumentElement!.Remove());
        Assert.Null(edited.Document.DocumentElement);
        Assert.Empty(edited.ToMarkdown().Trim());
        Assert.NotNull(edited.SemanticDocument);
        Assert.Empty(edited.CreateDocumentForConversion().Body!.TextContent);
        Assert.Contains("Before", original.ToMarkdown());
    }

    [Theory]
    [InlineData(HtmlElement.HtmlNamespace)]
    [InlineData("http://www.w3.org/2000/svg")]
    public void ReflectedIdentityIgnoresNamespacedAttributesWhileQualifiedAccessRetainsThem(string elementNamespace) {
        var document = new HtmlDocument(AngleSharpDomServices.Instance, "authored");
        HtmlElement element = document.CreateElement("g", elementNamespace);
        document.AppendChild(element);
        element.SetAttribute("id", "foreign-id", "urn:test");
        element.SetAttribute("class", "foreign-class", "urn:test");
        element.SetAttribute("id", "ordinary-id");
        element.SetAttribute("class", "ordinary-class");
        Assert.Equal("ordinary-id", element.Id);
        Assert.Equal("ordinary-class", element.ClassName);
        Assert.Equal(new[] { "ordinary-class" }, element.ClassList);
        Assert.Equal("foreign-id", element.GetAttribute("id"));
        Assert.Equal("foreign-class", element.GetAttribute("class"));
        Assert.Same(element, document.QuerySelector("#ordinary-id"));
        Assert.Same(element, document.QuerySelector(".ordinary-class"));
        Assert.Null(document.QuerySelector("#foreign-id"));
        element.RemoveAttribute("id");
        Assert.Null(element.GetAttribute("urn:test", "id"));
        Assert.Equal("ordinary-id", element.Id);
        element.RemoveAttribute(string.Empty, "class");
        Assert.True(element.HasAttribute("class"));
        Assert.Empty(element.ClassName);
        Assert.Empty(element.ClassList);
        Assert.Null(document.QuerySelector(".foreign-class"));
    }

    [Theory]
    [InlineData("<div><span>x</span></div>", 4, null)]
    [InlineData("<div><span>x</span></div>", null, 3)]
    [InlineData("<template><span>x</span></template>", 5, null)]
    public void AlternateParserReportsTheSameSharedNodeAndDepthFailures(string html, int? nodes, int? depth) {
        var options = new HtmlConversionDocumentOptions { Limits = new HtmlConversionLimits { MaxHtmlNodes = nodes, MaxHtmlDepth = depth } };
        HtmlDomLimitException expected = Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.Parse(html, options));
        options.ParserProvider = new ForwardingParser();
        HtmlDomLimitException actual = Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.Parse(html, options));
        Assert.Equal(expected.Code, actual.Code);
        Assert.Equal(expected.LimitSource, actual.LimitSource);
        Assert.Equal(expected.Limit, actual.Limit);
        Assert.IsType<HtmlParseLimitException>(actual.InnerException);
    }

    [Fact]
    public void ProviderInputFailuresRetainDetailsAndUnknownBudgetsRemainProviderSpecific() {
        var failure = new HtmlParseLimitException(nameof(HtmlParseOptions.MaxInputCharacters), 12, 10);
        var options = new HtmlConversionDocumentOptions { ParserProvider = new FailingParser(failure) };
        HtmlDomLimitException translated = Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.Parse("x", options));
        Assert.Equal(HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded, translated.Code);
        Assert.Equal("MaxInputCharacters", translated.LimitSource);
        Assert.Equal(12, translated.Actual);
        Assert.Equal(10, translated.Limit);
        Assert.Same(failure, translated.InnerException);
        var custom = new HtmlParseLimitException("ProviderWorkUnits", 12, 10);
        options.ParserProvider = new FailingParser(custom);
        Assert.Same(custom, Assert.Throws<HtmlParseLimitException>(() => HtmlConversionDocument.Parse("x", options)));
    }

    [Fact]
    public async Task ConcurrentFirstProjectionPublishesOneCompleteStatePerSnapshot() {
        HtmlDocument[] documents = Enumerable.Range(0, 8).Select(index => {
            HtmlDocument document = AngleSharpHtmlParser.Instance.Parse("<p>Attached</p>", new HtmlParseOptions()).Clone();
            document.QuerySelector("p")!.SetAttribute("id", "document-" + index);
            for (int child = 0; child < 100; child++) document.Body!.AppendChild(document.CreateElement("span")).TextContent = "item";
            return document.Freeze();
        }).ToArray();
        using var start = new ManualResetEventSlim();
        Task<NativeDomBridge.NativeState>[] workers = Enumerable.Range(0, 24).Select(index => Task.Run(() => {
            start.Wait();
            return NativeDomBridge.GetState(documents[index % documents.Length]);
        })).ToArray();
        start.Set();
        NativeDomBridge.NativeState[] results = await Task.WhenAll(workers);
        for (int index = 0; index < results.Length; index++) {
            HtmlDocument document = documents[index % documents.Length];
            Assert.Same(results[index % documents.Length], results[index]);
            Assert.Equal("document-" + index % documents.Length, results[index].Native.QuerySelector("p")!.Id);
            Assert.Equal(results[index].ToNative.Count, results[index].ToOwned.Count);
            Assert.Same(document, results[index].ToOwned[results[index].Native]);
            foreach (var pair in results[index].ToNative) Assert.Same(document.GetNode(pair.Key), results[index].ToOwned[pair.Value]);
        }
    }

    [Fact]
    public async Task ConcurrentDetachedProjectionAndCallbackSnapshotsRetainNodeIdentity() {
        HtmlDocument document = AngleSharpHtmlParser.Instance.Parse("<p>Attached</p>", new HtmlParseOptions()).Clone();
        HtmlElement[] detached = Enumerable.Range(0, 12).Select(index => {
            HtmlElement root = document.CreateElement("div");
            root.TextContent = "Detached " + index;
            return root;
        }).ToArray();
        document.Freeze();
        var projected = await Task.WhenAll(Enumerable.Range(0, 48).Select(index => Task.Run(() => NativeDomBridge.GetNative(detached[index % detached.Length]))));
        for (int index = 0; index < projected.Length; index++) Assert.Same(projected[index % detached.Length], projected[index]);
        var native = NativeDomBridge.GetNativeDocument(document);
        var paragraph = native.QuerySelector("p")!;
        try {
            HtmlElement[] callbacks = await Task.WhenAll(Enumerable.Range(0, 24).Select(_ => Task.Run(() => NativeDomBridge.Wrap(paragraph))));
            foreach (HtmlElement callback in callbacks) {
                Assert.Same(callbacks[0], callback);
                Assert.Same(paragraph, NativeDomBridge.GetCallbackNative(callback, callback.Document));
            }
            NativeDomBridge.ReleaseCallbackSnapshot(native);
            Assert.NotSame(callbacks[0].Document, NativeDomBridge.Wrap(paragraph).Document);
            Assert.Same(paragraph, NativeDomBridge.GetCallbackNative(callbacks[0], callbacks[0].Document));
        } finally { NativeDomBridge.ReleaseCallbackSnapshot(native); }
    }

    private sealed class ForwardingParser : IHtmlParserProvider {
        public string Id => "forwarding";
        public HtmlDocument Parse(string source, HtmlParseOptions options, CancellationToken cancellationToken = default) => AngleSharpHtmlParser.Instance.Parse(source, options, cancellationToken);
    }
    private sealed class FailingParser : IHtmlParserProvider {
        private readonly HtmlParseLimitException _failure;
        public FailingParser(HtmlParseLimitException failure) => _failure = failure;
        public string Id => "failing";
        public HtmlDocument Parse(string source, HtmlParseOptions options, CancellationToken cancellationToken = default) => throw _failure;
    }
}
