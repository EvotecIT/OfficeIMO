using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Providers;
using OfficeIMO.Markdown.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlImportTests {
    private static HtmlDocument Parse(string html) => AngleSharpHtmlParser.Instance.Parse(html, new HtmlParseOptions());

    [Fact]
    public void ImportCreatesIndependentNodesAndKeepsDestinationIdentity() {
        HtmlDocument source = Parse("<section id='source'><p>Decoded &amp; text</p><!--kept--></section>");
        HtmlDocument target = Parse("<main>Existing</main>").Clone();
        HtmlElement original = source.QuerySelector("section")!;
        HtmlElement main = target.QuerySelector("main")!;
        Assert.Throws<ArgumentException>(() => main.AppendChild(original));
        HtmlNode imported = target.ImportNode(original);
        Assert.Null(imported.Parent);
        Assert.Same(target, imported.Document);
        Assert.Same(imported, target.GetNode(imported.NodeId));
        Assert.NotEqual(main.NodeId, imported.NodeId);
        Assert.NotNull(original.SourceIndex);
        Assert.All(imported.Descendants().Prepend(imported), node => Assert.Null(node.SourceIndex));
        main.AppendChild(imported);
        Assert.Equal("Decoded & text", imported.TextContent);
        Assert.Contains("<!--kept-->", imported.OuterHtml);
        imported.QuerySelector("p")!.TextContent = "Changed";
        Assert.Equal("Decoded & text", original.TextContent);
        Assert.Equal("ExistingChanged", main.TextContent);
    }

    [Fact]
    public void ImportPreservesForeignAttributesAndNestedTemplateContents() {
        HtmlDocument source = Parse("<template><svg viewBox='0 0 10 10'><use xlink:href='#shape'/></svg><template><b>Nested</b></template></template>");
        HtmlDocument target = Parse("<p>Target</p>").Clone();
        HtmlElement imported = (HtmlElement)target.ImportNode(source.QuerySelector("template")!);
        target.Body!.AppendChild(imported);
        Assert.Empty(imported.ChildNodes);
        HtmlElement svg = imported.TemplateContent!.QuerySelector("svg")!;
        Assert.Equal("http://www.w3.org/2000/svg", svg.NamespaceUri);
        Assert.Equal("0 0 10 10", svg.GetAttribute("viewBox"));
        Assert.Equal("#shape", svg.QuerySelector("use")!.GetAttribute("http://www.w3.org/1999/xlink", "href"));
        HtmlElement nested = imported.TemplateContent.QuerySelector("template")!;
        Assert.Equal("Nested", nested.TemplateContent!.TextContent);
        Assert.Null(target.QuerySelector("b"));
        Assert.Equal(source.QuerySelector("template")!.OuterHtml, imported.OuterHtml);
    }

    [Fact]
    public void ImportedFragmentSplicesChildrenWithoutMovingSourceContent() {
        HtmlDocument source = Parse("<template>Lead<b>Bold</b>Tail</template>");
        HtmlNode original = source.QuerySelector("template")!.TemplateContent!;
        HtmlDocument target = Parse("<main></main>").Clone();
        HtmlNode fragment = target.ImportNode(original);
        Assert.Equal(HtmlNodeKind.DocumentFragment, fragment.Kind);
        target.QuerySelector("main")!.AppendChild(fragment);
        Assert.Empty(fragment.ChildNodes);
        Assert.Equal(3, original.ChildNodes.Count);
        Assert.Equal("Lead<b>Bold</b>Tail", target.QuerySelector("main")!.InnerHtml);
        Assert.Contains("Lead**Bold**Tail", HtmlConversionDocument.FromDocument(target).ToMarkdown());
    }

    [Fact]
    public void ShallowImportKeepsAttributesButOmitsAllContent() {
        HtmlDocument source = Parse("<section id='s'><b>Normal</b></section><template id='t'><b>Template</b></template>");
        HtmlDocument target = Parse("").Clone();
        HtmlElement section = (HtmlElement)target.ImportNode(source.QuerySelector("section")!, deep: false);
        HtmlElement template = (HtmlElement)target.ImportNode(source.QuerySelector("template")!, deep: false);
        Assert.Equal("s", section.Id);
        Assert.Empty(section.ChildNodes);
        Assert.Equal("t", template.Id);
        Assert.Empty(template.TemplateContent!.ChildNodes);
    }

    [Fact]
    public void ImportRejectsDocumentsAndCancelledRequestsBeforeMutation() {
        HtmlDocument source = Parse("<!doctype html><p>Source</p>");
        HtmlDocument target = Parse("").Clone();
        long revision = target.Revision;
        Assert.Throws<ArgumentException>(() => target.ImportNode(source));
        Assert.Throws<OperationCanceledException>(() => target.ImportNode(source.Body!, cancellationToken: new CancellationToken(true)));
        Assert.Equal(revision, target.Revision);
        var doctype = Assert.IsType<HtmlDocumentType>(target.ImportNode(source.FirstChild!));
        Assert.Equal("html", doctype.Name);
        target.Freeze();
        Assert.Throws<InvalidOperationException>(() => target.ImportNode(source.Body!));
    }

    [Fact]
    public void ImportingIntoAnEditLeavesBothOriginalSnapshotsIntact() {
        var source = HtmlConversionDocument.Parse("<p>Source</p>");
        var destination = HtmlConversionDocument.Parse("<main>Target</main>");
        var edited = destination.Edit(document => document.QuerySelector("main")!.AppendChild(document.ImportNode(source.Document.QuerySelector("p")!)));
        Assert.Equal("Source", source.Document.Body!.TextContent);
        Assert.Equal("Target", destination.Document.Body!.TextContent);
        Assert.Contains("Source", edited.ToMarkdown());
        Assert.True(edited.Document.IsReadOnly);
    }

    [Fact]
    public void SameDocumentImportCopiesBeforeInsertionWithoutReusingHandles() {
        HtmlDocument document = Parse("<section id='original'><b>Source</b></section>").Clone();
        HtmlElement original = document.QuerySelector("section")!;
        HtmlElement copied = (HtmlElement)document.ImportNode(original);
        Assert.Null(copied.Parent);
        Assert.NotEqual(original.NodeId, copied.NodeId);
        Assert.NotEqual(original.FirstChild!.NodeId, copied.FirstChild!.NodeId);
        copied.SetAttribute("id", "copy");
        copied.QuerySelector("b")!.TextContent = "Changed";
        document.Body!.AppendChild(copied);
        Assert.Equal("Source", document.QuerySelector("#original")!.TextContent);
        Assert.Equal("Changed", document.QuerySelector("#copy")!.TextContent);
        Assert.Equal(2, document.QuerySelectorAll("section").Count);
    }
}
