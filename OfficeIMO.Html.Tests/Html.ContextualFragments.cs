using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlContextualFragmentTests {
    [Fact]
    public void DefaultDocumentEngineReturnsOwnedSnapshotsWithoutProviderTypes() {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument("<main><p id='value'>Ready</p></main>");

        Assert.True(document.IsReadOnly);
        Assert.StartsWith("AngleSharp/", document.ProviderId, StringComparison.Ordinal);
        Assert.Equal("Ready", document.QuerySelector("#value")!.TextContent);
        Assert.Equal(typeof(HtmlDocument), document.GetType());
    }

    [Fact]
    public void TableFragmentUsesTheSuppliedRowContextAndCanBeImported() {
        HtmlDocument target = HtmlDocumentEngine.Default
            .ParseDocument("<table><tbody><tr id='row'><th>Head</th></tr></tbody></table>")
            .CloneAttached();
        HtmlElement row = target.QuerySelector("#row")!;

        HtmlDocumentFragment fragment = HtmlDocumentEngine.Default.ParseFragment(
            "<td>A</td><td>B</td>",
            row);

        Assert.True(fragment.Document.IsReadOnly);
        Assert.Equal(HtmlNodeKind.DocumentFragment, fragment.Kind);
        Assert.Equal(new[] { "td", "td" }, fragment.Children.Select(element => element.LocalName));
        Assert.Equal("AB", fragment.TextContent);
        Assert.Equal(0, fragment.Children.First().SourceIndex);
        Assert.Equal(target.Mode, fragment.Document.Mode);
        Assert.NotSame(target, fragment.Document);

        HtmlDocumentFragment imported = (HtmlDocumentFragment)target.ImportNode(fragment);
        row.AppendChild(imported);
        Assert.Equal(new[] { "th", "td", "td" }, row.Children.Select(element => element.LocalName));
        Assert.Empty(imported.ChildNodes);
    }

    [Fact]
    public void ForeignFragmentRetainsSvgAndMathMlNamespacesAndNames() {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(
            "<svg><g id='svg-context'></g></svg><math><mrow id='math-context'></mrow></math>");

        HtmlDocumentFragment svg = HtmlDocumentEngine.Default.ParseFragment(
            "<linearGradient id='paint'><stop offset='1'/></linearGradient>",
            document.QuerySelector("#svg-context")!);
        HtmlElement gradient = Assert.Single(svg.Children);
        Assert.Equal("http://www.w3.org/2000/svg", gradient.NamespaceUri);
        Assert.Equal("linearGradient", gradient.LocalName);
        Assert.Equal(0, gradient.SourceIndex);
        Assert.Equal("http://www.w3.org/2000/svg", Assert.Single(gradient.Children).NamespaceUri);

        HtmlDocumentFragment math = HtmlDocumentEngine.Default.ParseFragment(
            "<mi>x</mi>",
            document.QuerySelector("#math-context")!);
        HtmlElement identifier = Assert.Single(math.Children);
        Assert.Equal("http://www.w3.org/1998/Math/MathML", identifier.NamespaceUri);
        Assert.Equal("mi", identifier.LocalName);
    }

    [Fact]
    public void ForeignFragmentRetainsNodesReparentedByHtmlBreakoutTokens() {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(
            "<svg><g id='svg-context'></g></svg><math><mrow id='math-context'></mrow></math>");

        HtmlDocumentFragment svg = HtmlDocumentEngine.Default.ParseFragment(
            "<circle id='svg-child'/><div id='svg-breakout'>HTML</div><path id='svg-after'/>",
            document.QuerySelector("#svg-context")!);
        Assert.Equal("http://www.w3.org/2000/svg", svg.QuerySelector("#svg-child")!.NamespaceUri);
        Assert.Equal(HtmlElement.HtmlNamespace, svg.QuerySelector("#svg-breakout")!.NamespaceUri);
        Assert.NotNull(svg.QuerySelector("#svg-after"));
        Assert.Equal("HTML", svg.QuerySelector("#svg-breakout")!.TextContent);

        HtmlDocumentFragment math = HtmlDocumentEngine.Default.ParseFragment(
            "<mi id='math-child'>x</mi><div id='math-breakout'>HTML</div><mo id='math-after'>+</mo>",
            document.QuerySelector("#math-context")!);
        Assert.Equal("http://www.w3.org/1998/Math/MathML", math.QuerySelector("#math-child")!.NamespaceUri);
        Assert.Equal(HtmlElement.HtmlNamespace, math.QuerySelector("#math-breakout")!.NamespaceUri);
        Assert.NotNull(math.QuerySelector("#math-after"));
    }

    [Fact]
    public void FragmentRetainsAncestorFormParsingContext() {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(
            "<form id='owner'><div id='context'></div></form>");

        HtmlDocumentFragment fragment = HtmlDocumentEngine.Default.ParseFragment(
            "<form id='nested'><input name='value'></form>",
            document.QuerySelector("#context")!);

        Assert.Null(fragment.QuerySelector("#nested"));
        Assert.NotNull(fragment.QuerySelector("input[name='value']"));
    }

    [Fact]
    public void EnvelopeRetainsFosterParentedAndContextEscapingNodes() {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(
            "<form id='owner'><table id='table'></table><section id='section'></section></form>");

        HtmlDocumentFragment table = HtmlDocumentEngine.Default.ParseFragment(
            "before<tr><td id='cell'>Cell</td></tr>after",
            document.QuerySelector("#table")!);
        Assert.Contains("before", table.TextContent, StringComparison.Ordinal);
        Assert.Contains("after", table.TextContent, StringComparison.Ordinal);
        Assert.Equal("Cell", table.QuerySelector("#cell")!.TextContent);

        HtmlDocumentFragment escaped = HtmlDocumentEngine.Default.ParseFragment(
            "</section><p id='escaped'>Kept</p>",
            document.QuerySelector("#section")!);
        Assert.Equal("Kept", escaped.QuerySelector("#escaped")!.TextContent);
    }

    [Fact]
    public void EnvelopeEndsAtSourceEofForRcdataAndPlaintext() {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(
            "<form><textarea id='textarea'></textarea><div id='context'></div></form>");

        HtmlDocumentFragment rcdata = HtmlDocumentEngine.Default.ParseFragment(
            "<b>literal</b>",
            document.QuerySelector("#textarea")!);
        Assert.Equal("<b>literal</b>", rcdata.TextContent);
        Assert.Empty(rcdata.Children);

        HtmlDocumentFragment plaintext = HtmlDocumentEngine.Default.ParseFragment(
            "<plaintext id='plain'>literal",
            document.QuerySelector("#context")!);
        Assert.Equal("literal", plaintext.QuerySelector("#plain")!.TextContent);
    }

    [Fact]
    public void TemplateContextBelowFormReturnsTemplateContent() {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(
            "<form><template id='context'></template></form>");

        HtmlDocumentFragment fragment = HtmlDocumentEngine.Default.ParseFragment(
            "<span id='inside'>Template value</span>",
            document.QuerySelector("#context")!);

        Assert.Equal("Template value", fragment.QuerySelector("#inside")!.TextContent);
    }

    [Fact]
    public void ForeignContextBelowFormRetainsBothParsingContexts() {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(
            "<form id='owner'><svg><g id='context'></g></svg></form>");

        HtmlDocumentFragment fragment = HtmlDocumentEngine.Default.ParseFragment(
            "<div id='breakout'></div><form id='nested'><input id='kept'></form>",
            document.QuerySelector("#context")!);

        Assert.NotNull(fragment.QuerySelector("#breakout"));
        Assert.Null(fragment.QuerySelector("#nested"));
        Assert.NotNull(fragment.QuerySelector("#kept"));
    }

    [Fact]
    public void FragmentParsingHonorsSourceTreeAndCancellationLimits() {
        HtmlElement context = HtmlDocumentEngine.Default.ParseDocument("<div id='context'></div>")
            .QuerySelector("#context")!;

        Assert.Throws<HtmlParseLimitException>(() => HtmlDocumentEngine.Default.ParseFragment(
            "<span><b>value</b></span>",
            context,
            new HtmlParseOptions { MaxNodes = 2 }));
        Assert.Throws<HtmlParseLimitException>(() => HtmlDocumentEngine.Default.ParseFragment(
            "<span><b>value</b></span>",
            context,
            new HtmlParseOptions { MaxDepth = 1 }));

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => HtmlDocumentEngine.Default.ParseFragment(
            "<span>value</span>",
            context,
            cancellationToken: cancellation.Token));
    }
}
