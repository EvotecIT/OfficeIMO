using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Providers;
using OfficeIMO.Markdown.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlFoundationLimitTests {
    private const string TemplateSource = "<!doctype html><template><div><span>x</span></div></template>";
    private static HtmlDocument Parse(string source) => AngleSharpHtmlParser.Instance.Parse(source, new HtmlParseOptions());
    private static HtmlConversionDocumentOptions Options(int? nodes = null, int? depth = null, int? characters = null) =>
        new HtmlConversionDocumentOptions { Limits = new HtmlConversionLimits { MaxHtmlNodes = nodes, MaxHtmlDepth = depth, MaxInputCharacters = characters } };

    [Theory]
    [InlineData(8, 5, "MaxHtmlNodes")]
    [InlineData(9, 4, "MaxHtmlDepth")]
    public void TemplateBudgetsAgreeForNativeOwnedAndConversionInputs(int nodes, int depth, string limit) {
        var options = Options(nodes, depth);
        Assert.Equal(limit, Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.Parse(TemplateSource, options)).LimitSource);
        Assert.Equal(limit, Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.FromDocument(Parse(TemplateSource), options)).LimitSource);
        Assert.Throws<HtmlParseLimitException>(() => AngleSharpHtmlParser.Instance.Parse(TemplateSource, new HtmlParseOptions { MaxNodes = nodes, MaxDepth = depth }));
    }

    [Fact]
    public void TemplateBudgetIncludesFragmentWithoutAddingElementDepth() {
        var options = Options(9, 5);
        HtmlConversionDocument native = HtmlConversionDocument.Parse(TemplateSource, options);
        HtmlDocument parsed = AngleSharpHtmlParser.Instance.Parse(TemplateSource, new HtmlParseOptions { MaxNodes = 9, MaxDepth = 5 });
        HtmlConversionDocument captured = HtmlConversionDocument.FromDocument(parsed, options);
        Assert.Equal(parsed.OuterHtml, native.Document.OuterHtml);
        Assert.Equal(parsed.OuterHtml, captured.Document.OuterHtml);
        Assert.Equal("x", native.Document.QuerySelector("template")!.TemplateContent!.QuerySelector("span")!.TextContent);
    }

    [Fact]
    public void NativeTemplateContentsRetainStylesheetAndSemanticBudgets() {
        var css = Options();
        css.Limits.MaxCssBytes = 4;
        Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.Parse("<template><style>p{color:red}</style></template>", css));
        var metadata = Options();
        metadata.Limits.MaxSemanticMetadataCharacters = 4;
        Assert.Equal("MaxSemanticMetadataCharacters", Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.Parse("<template><p data-officeimo-kind='oversized'>x</p></template>", metadata)).LimitSource);
        Assert.Equal("MaxHtmlDepth", Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.Parse("<template><iframe srcdoc='<div><span>x</span></div>'></iframe></template>", Options(depth: 6))).LimitSource);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConversionCapturesAttachedNodesWithoutRetainingDetachedEditingHistory(bool freeze) {
        HtmlDocument original = Parse("<p>Kept</p><template><b>Template</b></template>").Clone();
        HtmlElement p = original.QuerySelector("p")!;
        HtmlNode content = original.QuerySelector("template")!.TemplateContent!;
        HtmlElement detached = original.CreateElement("invalid@detached", "urn:foreign");
        detached.TextContent = new string('x', 4096);
        if (freeze) original.Freeze();
        HtmlConversionDocument captured = HtmlConversionDocument.FromDocument(original, Options(nodes: 10, characters: 256));
        Assert.NotSame(original, captured.Document);
        Assert.Null(captured.Document.GetNode(detached.NodeId));
        Assert.NotNull(original.Clone().GetNode(detached.NodeId));
        Assert.Equal(p.SourceIndex, captured.Document.GetNode(p.NodeId)!.SourceIndex);
        Assert.Equal("Kept", captured.Document.GetNode(p.NodeId)!.TextContent);
        Assert.Equal("Template", captured.Document.GetNode(content.NodeId)!.TextContent);
        Assert.Contains("Kept", captured.ToMarkdown());
    }

    [Fact]
    public void AlternateParserCaptureAndConversionEditsDiscardDetachedNodes() {
        HtmlDocument supplied = Parse("<p>Kept</p>").Clone();
        HtmlNode detached = supplied.CreateTextNode(new string('x', 4096));
        var options = Options(nodes: 5, characters: 128);
        options.ParserProvider = new SuppliedParser(supplied);
        HtmlConversionDocument conversion = HtmlConversionDocument.Parse("source", options);
        Assert.Null(conversion.Document.GetNode(detached.NodeId));
        int removedId = 0;
        HtmlDocument? editor = null;
        HtmlConversionDocument edited = conversion.Edit(document => {
            editor = document;
            removedId = document.CreateTextNode(new string('x', 4096)).NodeId;
            document.Body!.TextContent = "Edited";
        });
        Assert.Null(edited.Document.GetNode(removedId));
        Assert.True(editor!.IsReadOnly);
        Assert.Contains("Edited", edited.ToMarkdown());
        Assert.Contains("Kept", conversion.ToMarkdown());
        Assert.NotNull(supplied.GetNode(detached.NodeId));
    }

    [Theory]
    [InlineData("text")]
    [InlineData("attribute")]
    [InlineData("doctype")]
    public void OwnedDataLimitRunsBeforeNativeProjection(string kind) {
        var owned = new HtmlDocument(AngleSharpDomServices.Instance, "authored");
        HtmlElement root = owned.CreateElement("invalid@foreign", "urn:foreign");
        string value = new string('x', 1024);
        if (kind == "doctype") owned.AppendChild(owned.CreateDocumentType("html", value));
        owned.AppendChild(root);
        if (kind == "text") root.TextContent = value;
        if (kind == "attribute") root.SetAttribute("title", value);
        Assert.Equal("MaxInputCharacters", Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.FromDocument(owned, Options(characters: 128))).LimitSource);
    }

    [Fact]
    public void OwnedDataLimitAggregatesSmallValuesBeforeNativeProjection() {
        var owned = new HtmlDocument(AngleSharpDomServices.Instance, "authored");
        HtmlElement root = owned.CreateElement("invalid@foreign", "urn:foreign");
        owned.AppendChild(root);
        for (int i = 0; i < 20; i++) root.AppendChild(owned.CreateTextNode("1234567890"));
        Assert.Equal("MaxInputCharacters", Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.FromDocument(owned, Options(characters: 128))).LimitSource);
    }

    [Theory]
    [InlineData("&")]
    [InlineData("\u00a0")]
    [InlineData("\"")]
    public void CanonicalSerializationHonorsExactExpandedLength(string character) {
        HtmlDocument owned = Parse("<p></p>").Edit(document => {
            document.QuerySelector("p")!.TextContent = string.Concat(Enumerable.Repeat(character, 12));
            document.QuerySelector("p")!.SetAttribute("title", string.Concat(Enumerable.Repeat(character, 12)));
        });
        string expected = owned.OuterHtml;
        Assert.Equal(expected, HtmlConversionDocument.FromDocument(owned, Options(characters: expected.Length)).SourceHtml);
        Assert.Equal("MaxInputCharacters", Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.FromDocument(owned, Options(characters: expected.Length - 1))).LimitSource);
    }

    [Fact]
    public void CancellationIsObservedAtNativeValidationAndProjectionBoundaries() {
        var cancelled = new CancellationToken(true);
        var native = HtmlDocumentParser.ParseDocument("<template><iframe srcdoc='<p>Nested</p>'></iframe></template>");
        Assert.Throws<OperationCanceledException>(() => HtmlConversionInputGuard.ValidateDocument(native, new HtmlConversionLimits(), cancelled));
        Assert.Throws<OperationCanceledException>(() => NativeDomBridge.Import(native, cancellationToken: cancelled));

        HtmlDocument owned = Parse("<p>Before</p>").Edit(document => document.QuerySelector("p")!.TextContent = "Changed");
        Assert.Throws<OperationCanceledException>(() => NativeDomBridge.GetNativeDocument(owned, cancelled));
        var projected = NativeDomBridge.GetNativeDocument(owned);
        Assert.Equal("Changed", projected.Body!.TextContent);
        Assert.Throws<OperationCanceledException>(() => NativeDomBridge.GetNativeDocument(owned, cancelled));
        Assert.Same(projected, NativeDomBridge.GetNativeDocument(owned));
    }

    private sealed class SuppliedParser : IHtmlParserProvider {
        private readonly HtmlDocument _document;
        internal SuppliedParser(HtmlDocument document) { _document = document; }
        public string Id => "supplied";
        public HtmlDocument Parse(string source, HtmlParseOptions options, CancellationToken cancellationToken = default) => _document;
    }
}
