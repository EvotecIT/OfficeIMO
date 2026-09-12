using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlFoundationTests {
    private static HtmlDocument Parse(string source) => AngleSharpHtmlParser.Instance.Parse(source, new HtmlParseOptions());

    [Fact]
    public void RecoveredDoctypeSurvivesEditingAndSerialization() {
        HtmlDocument original = Parse("<!DOCTYPE odd@name><p>Before</p>");
        HtmlDocument edited = original.Edit(document => document.QuerySelector("p")!.TextContent = "After");
        Assert.Contains("<!DOCTYPE odd@name>", edited.OuterHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("odd@name", Assert.IsType<HtmlDocumentType>(edited.FirstChild).Name);
        Assert.Contains("After", HtmlConversionDocument.FromDocument(edited).ToMarkdown());
    }

    [Fact]
    public void AuthoredHtmlNamesAgreeAcrossQueriesStylesAndConversion() {
        HtmlDocument original = Parse("<p>Before</p>");
        HtmlDocument edited = original.Edit(document => {
            HtmlElement p = document.QuerySelector("p")!;
            p.SetAttribute("ID", "changed"); p.SetAttribute("CLASS", "notice"); p.SetAttribute("STYLE", "color:#123456");
            HtmlElement link = document.CreateElement("A"); link.SetAttribute("HREF", "https://example.org/guide"); link.TextContent = "Guide"; p.AppendChild(link);
        });
        HtmlElement changed = edited.QuerySelector("#changed.notice")!;
        Assert.NotNull(changed);
        Assert.Equal("id", changed.Attributes[0].Name);
        Assert.Equal("a", edited.QuerySelector("a")!.LocalName);
        Assert.Contains("https://example.org/guide", HtmlConversionDocument.FromDocument(edited).ToMarkdown());
        Assert.Contains("#123456", HtmlComputedStyleEngine.Compute(edited)[changed].GetValue("color"));
    }

    [Fact]
    public void ForeignAndNamespacedAttributeNamesRetainCaseAndIdentity() {
        HtmlDocument document = Parse("<p></p><svg viewBox='0 0 10 10'></svg>").Clone();
        HtmlElement p = document.QuerySelector("p")!, svg = document.QuerySelector("svg")!;
        p.SetAttribute("xml:lang", "plain");
        Assert.Equal("xml:lang", p.Attributes[0].LocalName);
        Assert.Empty(p.Attributes[0].Prefix);
        p.SetAttribute("ID", "namespaced", "urn:test");
        p.SetAttribute("ID", "ordinary");
        Assert.Equal("ordinary", p.GetAttribute("ID"));
        Assert.Equal("namespaced", p.GetAttribute("urn:test", "ID"));
        p.RemoveAttribute("ID");
        Assert.Null(p.GetAttribute("id"));
        Assert.Equal("namespaced", p.GetAttribute("urn:test", "ID"));
        svg.SetAttribute("viewBox", "0 0 20 20");
        Assert.Null(svg.GetAttribute("viewbox"));
        Assert.Contains("viewBox=\"0 0 20 20\"", document.OuterHtml);
        Assert.Equal("ID", p.Attributes[1].LocalName);
    }

    [Fact]
    public void TolerantParserNamesSurviveAnUnrelatedEdit() {
        HtmlDocument document = Parse("<p>Before</p><odd@name strange@attribute='value'>Retained</odd@name>");
        HtmlDocument changed = document.Edit(edit => edit.QuerySelector("p")!.TextContent = "After");
        Assert.Contains("<odd@name strange@attribute=\"value\">Retained</odd@name>", changed.OuterHtml);
    }

    [Fact]
    public async Task FrozenSnapshotSupportsConcurrentQueriesAndSerialization() {
        HtmlDocument document = Parse("<section><p>One</p><p>Two</p></section>").Edit(edit => edit.QuerySelector("p")!.TextContent = "Edited");
        string serialized = document.OuterHtml;
        await Task.WhenAll(Enumerable.Range(0, 16).Select(_ => Task.Run(() => {
            for (int iteration = 0; iteration < 20; iteration++) {
                Assert.Equal(2, document.QuerySelectorAll("section > p").Count);
                Assert.Equal(serialized, document.OuterHtml);
            }
        })));
    }

    [Fact]
    public void Snapshot_EditPreservesIdentityWithoutMutatingSourceOrLeakedHandles() {
        var original = HtmlConversionDocument.Parse("<h1 id='title'>Before</h1><p>Body</p>");
        HtmlElement first = original.Document.QuerySelector("h1")!;
        HtmlDocument? editor = null;
        var changed = original.Edit(document => { editor = document; document.GetNode(first.NodeId)!.TextContent = "After"; });
        Assert.Equal("Before", first.TextContent);
        Assert.Equal("After", changed.Document.GetNode(first.NodeId)!.TextContent);
        Assert.NotEqual(original.Document.SnapshotId, changed.Document.SnapshotId);
        Assert.Same(changed.Document.GetNode(first.NodeId), changed.Document.QuerySelector("#title"));
        Assert.Throws<InvalidOperationException>(() => first.SetAttribute("x", "y"));
        Assert.Throws<InvalidOperationException>(() => editor!.CreateElement("div"));
        Assert.Contains("After", changed.ToMarkdown());
        Assert.DoesNotContain("After", original.ToMarkdown());
    }

    [Fact]
    public void FailedEditFreezesEscapedEditorAndLeavesSourceIntact() {
        HtmlDocument source = Parse("<p>Before</p>");
        HtmlDocument? leaked = null;
        Assert.Throws<FormatException>(() => source.Edit(edit => { leaked = edit; edit.Body!.TextContent = "Changed"; throw new FormatException(); }));
        Assert.Equal("Before", source.Body!.TextContent);
        Assert.Throws<InvalidOperationException>(() => leaked!.Body!.TextContent = "Again");
    }

    [Fact]
    public void DetachedNodesRemainQueryableSerializableAndClonable() {
        HtmlDocument document = Parse("<section><b>Old</b></section>").Clone();
        var section = document.QuerySelector("section")!;
        section.Remove();
        Assert.Equal("<section><b>Old</b></section>", section.OuterHtml);
        Assert.Equal("Old", section.QuerySelector("b")!.TextContent);
        var fresh = document.CreateElement("em");
        Assert.Equal("<em></em>", fresh.OuterHtml);
        fresh.TextContent = "New";
        section.AppendChild(fresh);
        Assert.True(fresh.Matches("section > em"));
        HtmlDocument clone = document.Clone();
        Assert.Equal(section.OuterHtml, clone.GetNode(section.NodeId)!.OuterHtml);
        Assert.Null(clone.GetNode(section.NodeId)!.Parent);
        Assert.Null(document.QuerySelector("section"));
    }

    [Fact]
    public void TemplatesRemainSeparateAndRejectCycles() {
        HtmlDocument document = Parse("<template id='t'><table><tr><td>Value</td></tr></table></template>").Clone();
        HtmlElement template = document.QuerySelector("template")!;
        Assert.Empty(template.ChildNodes);
        Assert.Null(document.QuerySelector("td"));
        Assert.Equal("Value", template.TemplateContent!.QuerySelector("td")!.TextContent);
        Assert.Contains("<tbody>", template.InnerHtml);
        Assert.Throws<ArgumentException>(() => template.TemplateContent.AppendChild(template));
        Assert.Throws<ArgumentException>(() => document.Body!.AppendChild(template.TemplateContent));
        template.TemplateContent.QuerySelector("td")!.TextContent = "Edited";
        Assert.Contains("Edited", document.Clone().OuterHtml);
    }

    [Fact]
    public void FragmentsSpliceChildrenAndInvalidMovesAreAtomic() {
        HtmlDocument document = Parse("<p>Kept</p>").Clone();
        HtmlNode fragment = document.CreateFragment();
        fragment.AppendChild(document.CreateElement("b"));
        fragment.AppendChild(document.CreateTextNode("Tail"));
        Assert.Throws<ArgumentException>(() => document.AppendChild(fragment));
        Assert.Equal(2, fragment.ChildNodes.Count);
        document.Body!.AppendChild(fragment);
        Assert.Empty(fragment.ChildNodes);
        Assert.Equal("<p>Kept</p><b></b>Tail", document.Body.InnerHtml);
        Assert.Throws<ArgumentException>(() => document.Body.AppendChild(document.DocumentElement!));
        Assert.Throws<ArgumentException>(() => document.Body.AppendChild(Parse("<p>Other</p>").Body!));
    }

    [Theory]
    [InlineData("<!doctype html><p>x</p>", HtmlDocumentMode.Standards)]
    [InlineData("<p>x</p>", HtmlDocumentMode.Quirks)]
    [InlineData("<!DOCTYPE HTML PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'><p>x</p>", HtmlDocumentMode.LimitedQuirks)]
    public void DocumentModeSurvivesStructuralConversion(string html, HtmlDocumentMode mode) {
        HtmlDocument original = Parse(html);
        Assert.Equal(mode, original.Mode);
        var changed = HtmlConversionDocument.FromDocument(original.Edit(edit => edit.QuerySelector("p")!.TextContent = "y"));
        Assert.Equal(mode, changed.Document.Mode);
        Assert.Equal(mode, NativeDomBridge.Import(changed.CreateSourceDocumentForConversion()).Mode);
    }

    [Fact]
    public void ForeignAttributesAndSourcePositionsSurviveEdits() {
        HtmlDocument document = Parse("<!doctype html><svg viewBox='0 0 20 20'><use xmlns:xlink='http://www.w3.org/1999/xlink' xlink:href='#a'/></svg>");
        HtmlElement use = document.QuerySelector("use")!;
        Assert.Equal("#a", use.GetAttribute("http://www.w3.org/1999/xlink", "href"));
        var edit = document.Edit(tree => tree.QuerySelector("use")!.SetAttribute("xlink:href", "#b", "http://www.w3.org/1999/xlink"));
        Assert.Equal(use.SourceIndex, edit.GetNode(use.NodeId)!.SourceIndex);
        Assert.Contains("viewBox=", edit.OuterHtml);
        Assert.Contains("xlink:href=\"#b\"", edit.OuterHtml);
        Assert.Single(edit.QuerySelector("use")!.Attributes, a => a.LocalName == "href");
    }

    [Fact]
    public void ParserRejectsConfiguredLimitsAndCancellation() {
        var parser = AngleSharpHtmlParser.Instance;
        Assert.Throws<HtmlParseLimitException>(() => parser.Parse("12345", new HtmlParseOptions { MaxInputCharacters = 4 }));
        Assert.Throws<HtmlParseLimitException>(() => parser.Parse("<template><p>A</p><p>B</p></template>", new HtmlParseOptions { MaxNodes = 5 }));
        Assert.Throws<HtmlParseLimitException>(() => parser.Parse("<div><div><div>x</div></div></div>", new HtmlParseOptions { MaxDepth = 3 }));
        Assert.Throws<OperationCanceledException>(() => parser.Parse("<p>x</p>", new HtmlParseOptions(), new CancellationToken(true)));
        var options = new HtmlConversionDocumentOptions { Limits = new HtmlConversionLimits { MaxHtmlNodes = 5 } };
        Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.FromDocument(Parse("<template><p>A</p><p>B</p></template>"), options));
    }

    [Fact]
    public void AlternativeParserSuppliesOwnedNodesWithoutNativeTypes() {
        var provider = new AuthoredParser();
        var converted = HtmlConversionDocument.Parse("input", new HtmlConversionDocumentOptions { ParserProvider = provider });
        Assert.Equal(1, provider.Calls);
        Assert.Equal("authored", converted.Document.ProviderId);
        Assert.True(converted.Document.IsReadOnly);
        Assert.Contains("From provider", converted.ToMarkdown());
        provider.Returned!.Body!.TextContent = "Later mutation";
        Assert.Equal("From provider", converted.Document.Body!.TextContent);
    }

    [Fact]
    public async Task EncodingProviderIsScopedAndStreamPositionIsPreserved() {
        var provider = new CustomEncoding();
        byte[] bytes = Encoding.UTF8.GetBytes("<meta charset='custom-test'><p>Zażółć €</p>");
        using var stream = new MemoryStream(bytes);
        stream.Position = 7;
        var options = new HtmlConversionDocumentOptions { InputEncodingProvider = provider };
        var document = HtmlConversionDocument.Load(stream, options);
        Assert.Equal(7, stream.Position);
        Assert.Contains("Zażółć €", document.Document.Body!.TextContent);
        Assert.True(provider.Calls > 0);
        var asyncDocument = await HtmlConversionDocument.LoadAsync(stream, options);
        Assert.Equal(7, stream.Position);
        Assert.Equal(document.Document.OuterHtml, asyncDocument.Document.OuterHtml);
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void MarkdownFiltersReceiveStableReadOnlySnapshots() {
        HtmlElement? retained = null;
        var options = new HtmlToMarkdownOptions();
        options.ElementFilters.Add(element => {
            retained ??= element;
            Assert.True(element.Document.IsReadOnly);
            Assert.Throws<InvalidOperationException>(() => element.Remove());
            return element.Id == "remove";
        });
        var document = HtmlConversionDocument.Parse("<p id='remove'>Gone</p><p>Kept</p>");
        string markdown = document.ToMarkdown(options);
        Assert.DoesNotContain("Gone", markdown);
        Assert.Contains("Kept", markdown);
        Assert.Contains("Gone", retained!.Document.TextContent);
        Assert.Contains("Gone", retained.Document.OuterHtml);
    }

    [Fact]
    public void OwnedDocumentFeedsStylesMarkdownPngAndPagedPdf() {
        var document = HtmlConversionDocument.Parse("<!doctype html><style>h1{color:#234567} .next{break-before:page}</style><h1 id='title'>Foundation</h1><p>First page.</p><h1 class='next'>Second page</h1>");
        document = document.Edit(edit => edit.QuerySelector("#title")!.TextContent = "Edited foundation");
        var styles = HtmlComputedStyleEngine.Compute(document);
        Assert.Contains(document.Document.QuerySelector("#title")!, styles.Keys);
        Assert.Contains("Edited foundation", document.ToMarkdown());
        byte[] png = document.ToPng();
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Take(8).ToArray());
        byte[] pdf = document.ToPdfBytes(new HtmlToPdfOptions());
        var read = OfficeIMO.Pdf.PdfReadDocument.Open(pdf);
        Assert.Contains("Edited foundation", read.ExtractText());
        Assert.Contains("Second page", read.ExtractText());
        Assert.True(read.Pages.Count >= 2);
    }

    [Fact]
    public void PublicHtmlAndCallbackContractsDoNotExposeNativeDomTypes() {
        Assembly[] assemblies = { typeof(HtmlDocument).Assembly, typeof(HtmlConversionDocument).Assembly, typeof(HtmlInlineElementConversionContext).Assembly, typeof(AngleSharpHtmlParser).Assembly };
        foreach (Type type in assemblies.SelectMany(assembly => assembly.GetExportedTypes())) {
            foreach (MethodInfo method in type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly)) {
                AssertOwnedSignature(method.ReturnType, type.FullName + "." + method.Name);
                foreach (ParameterInfo parameter in method.GetParameters()) AssertOwnedSignature(parameter.ParameterType, type.FullName + "." + method.Name);
            }
            foreach (ConstructorInfo constructor in type.GetConstructors()) foreach (ParameterInfo parameter in constructor.GetParameters()) AssertOwnedSignature(parameter.ParameterType, type.FullName);
        }
        Assert.DoesNotContain(typeof(HtmlDocument).Assembly.GetReferencedAssemblies(), name => name.Name!.StartsWith("AngleSharp", StringComparison.Ordinal) || name.Name == "OfficeIMO.Core");
    }

    private static void AssertOwnedSignature(Type type, string? member = null) {
        Assert.False(type.Namespace?.StartsWith("AngleSharp", StringComparison.Ordinal) == true, member + " " + type.FullName);
        if (type.HasElementType) AssertOwnedSignature(type.GetElementType()!, member);
        foreach (Type argument in type.GetGenericArguments()) AssertOwnedSignature(argument, member);
    }

    private sealed class AuthoredParser : IHtmlParserProvider {
        public string Id => "authored";
        public int Calls { get; private set; }
        public HtmlDocument? Returned { get; private set; }
        public HtmlDocument Parse(string source, HtmlParseOptions options, CancellationToken cancellationToken = default) {
            Calls++;
            var document = new HtmlDocument(AngleSharpDomServices.Instance, Id);
            HtmlElement html = document.CreateElement("html"), body = document.CreateElement("body"), p = document.CreateElement("p");
            document.AppendChild(html); html.AppendChild(body); body.AppendChild(p); p.TextContent = "From provider";
            return Returned = document;
        }
    }

    private sealed class CustomEncoding : IHtmlEncodingProvider {
        public int Calls { get; private set; }
        public Encoding? ResolveLabel(string label) { Calls++; return label == "custom-test" ? Encoding.UTF8 : null; }
        public Encoding? ResolveMetaContent(string content) => null;
    }
}
