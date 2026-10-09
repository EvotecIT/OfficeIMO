using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixCollateralXhtmlTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixCollateralText Item(string fragment) => new() { Type = BookOnixTextType.Description,
        Audiences = [BookOnixContentAudience.EndCustomers], Texts = [new(fragment, "eng") { Format = BookOnixCollateralTextFormat.Xhtml }] };

    [Fact]
    public void ExplicitFragmentsPreserveSemanticMarkupAndNormalizeNamespaces() {
        string fragment = "<p xmlns='http://www.w3.org/1999/xhtml' xml:lang='en'>A <strong>bold</strong> &amp; <em>emphasized</em> description.</p>" +
            "<ol start='3' type='i'><li>First</li><li><a href='https://example.org/?a=1&amp;b=2'>Link</a></li></ol>" +
            "<p xmlns=''>Namespace-free paragraph.</p>";
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        var baseline = project.ExportOnix(BookOnixTests.Options(), BookOnixTests.TestSchema());
        var result = project.ExportOnix(BookOnixTests.Options() with { CollateralTexts = [Item(fragment)] }, BookOnixTests.TestSchema());
        var text = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "Text").Single();
        Assert.Equal("05", (string?)text.Attribute("textformat"));
        Assert.Equal("eng", (string?)text.Attribute("language"));
        Assert.All(text.Descendants(), e => Assert.Equal(Ns, e.Name.Namespace));
        Assert.Equal(new[] { "p", "ol", "p" }, text.Elements().Select(e => e.Name.LocalName));
        Assert.Equal("en", (string?)text.Elements().First().Attribute("lang"));
        Assert.Null(text.Elements().First().Attribute(XNamespace.Xml + "lang"));
        Assert.Equal("bold", text.Descendants(Ns + "strong").Single().Value);
        Assert.Equal("https://example.org/?a=1&b=2", (string?)text.Descendants(Ns + "a").Single().Attribute("href"));
        Assert.Equal("3", (string?)text.Element(Ns + "ol")!.Attribute("start"));
        Assert.Equal(baseline.Publication.Bytes, result.Publication.Bytes);
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void MessageCompositionPreservesMixedContentWhitespace() {
        const string fragment = "<p><strong>First</strong> <em>second</em></p><pre>  <code>A</code>\n  <code>B</code>  </pre>";
        var schema = BookOnixTests.TestSchema();
        var product = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { CollateralTexts = [Item(fragment)] }, schema);
        var message = BookOnixMessage.Create([product], schema);
        var source = XDocument.Load(new MemoryStream(product.Bytes), LoadOptions.PreserveWhitespace).Descendants(Ns + "Text").Single();
        var composed = XDocument.Load(new MemoryStream(message.Bytes), LoadOptions.PreserveWhitespace).Descendants(Ns + "Text").Single();
        Assert.Equal("First second", composed.Element(Ns + "p")!.Value);
        Assert.Equal("  A\n  B  ", composed.Element(Ns + "pre")!.Value);
        Assert.True(XNode.DeepEquals(source, composed));
    }

    [Fact]
    public void RichShortDescriptionsCountDecodedTextAndNotMarkup() {
        string body = string.Concat(Enumerable.Repeat("&#x1F600;", 350));
        var item = Item("<p><strong>" + body + "</strong></p>") with { Type = BookOnixTextType.ShortDescription };
        var result = BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with { CollateralTexts = [item] }, BookOnixTests.TestSchema());
        Assert.Equal(string.Concat(Enumerable.Repeat("😀", 350)), XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "Text").Single().Value);
        Assert.Throws<ArgumentException>(() => BookOnixTests.Project().ExportOnix(BookOnixTests.Options() with {
            CollateralTexts = [item with { Texts = [new("<p>" + body + "X</p>") { Format = BookOnixCollateralTextFormat.Xhtml }] }]
        }, BookOnixTests.TestSchema()));
    }

    [Theory]
    [InlineData("<script>alert(1)</script>")]
    [InlineData("<p onclick='run()'>Text</p>")]
    [InlineData("<p style='color:red'>Text</p>")]
    [InlineData("<a href='javascript:run()'>Text</a>")]
    [InlineData("<a href='file:///private/file'>Text</a>")]
    [InlineData("<a href='https://user:password@example.org'>Text</a>")]
    [InlineData("<a href='#local'>Text</a>")]
    [InlineData("<p id='local'>Text</p>")]
    [InlineData("<p xmlns='urn:foreign'>Text</p>")]
    [InlineData("<svg xmlns='http://www.w3.org/2000/svg'><text>Text</text></svg>")]
    [InlineData("<p lang='en' xml:lang='pl'>Text</p>")]
    [InlineData("<!DOCTYPE p SYSTEM 'https://example.org/test.dtd'><p>Text</p>")]
    [InlineData("<?target data?><p>Text</p>")]
    [InlineData("<!-- note --><p>Text</p>")]
    [InlineData("<p>Unclosed")]
    [InlineData("<p>&nbsp;</p>")]
    [InlineData("<p><br/></p>")]
    public void UnsupportedOrMalformedFragmentsAreRejectedWithoutMutation(string fragment) {
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { CollateralTexts = [Item(fragment)] }, BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Fact]
    public void FragmentDepthNodeAndFormatBoundariesAreEnforced() {
        var project = BookOnixTests.Project();
        string allowed = string.Concat(Enumerable.Repeat("<div>", 32)) + "Text" + string.Concat(Enumerable.Repeat("</div>", 32));
        project.ExportOnix(BookOnixTests.Options() with { CollateralTexts = [Item(allowed)] }, BookOnixTests.TestSchema());
        foreach (string rejected in new[] { "<div>" + allowed + "</div>", "Text" + string.Concat(Enumerable.Repeat("<br/>", 4096)) })
            Assert.Throws<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { CollateralTexts = [Item(rejected)] }, BookOnixTests.TestSchema()));
        Assert.Throws<ArgumentException>(() => project.ExportOnix(BookOnixTests.Options() with { CollateralTexts = [Item("<p>Text</p>") with {
            SourceTitles = [new("<p>Source</p>") { Format = BookOnixCollateralTextFormat.Xhtml }] }] }, BookOnixTests.TestSchema()));
        Assert.Throws<ArgumentOutOfRangeException>(() => project.ExportOnix(BookOnixTests.Options() with { CollateralTexts = [Item("Text") with {
            Texts = [new("Text") { Format = (BookOnixCollateralTextFormat)99 }] }] }, BookOnixTests.TestSchema()));
    }
}
