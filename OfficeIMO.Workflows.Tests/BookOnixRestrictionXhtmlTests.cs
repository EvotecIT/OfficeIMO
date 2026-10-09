using System.Xml.Linq;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookOnixRestrictionXhtmlTests {
    private static readonly XNamespace Ns = BookProject.OnixNamespace;
    private static BookOnixExportOptions Options(string text, BookOnixCollateralTextFormat format = BookOnixCollateralTextFormat.Xhtml) => BookOnixTests.Options() with {
        Commercial = new() { SalesRights = [new(BookOnixSalesRightsKind.Exclusive, new() { Worldwide = true }) {
            Restrictions = [new(BookOnixSalesRestrictionKind.Unspecified) { Notes = [new(text, "eng") { Format = format }] }]
        }] }
    };

    [Fact]
    public void MarkupAndWhitespaceSurviveCompositionAndPublishingUpdate() {
        var result = BookOnixTests.Project().ExportOnix(Options("<p xmlns='http://www.w3.org/1999/xhtml'>Only <em>named</em> outlets &amp; partners.</p>"), BookOnixTests.TestSchema());
        var note = XDocument.Load(new MemoryStream(result.Bytes)).Descendants(Ns + "SalesRestrictionNote").Single();
        Assert.Equal("05", note.Attribute("textformat")!.Value);
        Assert.Equal("eng", note.Attribute("language")!.Value);
        Assert.Equal("Only named outlets & partners.", note.Value);
        Assert.Single(note.Descendants(Ns + "em"));
        Assert.Equal(result.Bytes, BookOnixMessage.Create([result], BookOnixTests.TestSchema()).Bytes);
        var update = BookOnixMessage.CreateBlockUpdates([new(result) { ReplaceBlocks = [BookOnixBlock.PublishingDetail] }], BookOnixTests.TestSchema());
        Assert.True(XNode.DeepEquals(note, XDocument.Load(new MemoryStream(update.Bytes)).Descendants(Ns + "SalesRestrictionNote").Single()));
    }

    [Fact]
    public void PlainTextRemainsLiteralAndDecodedBoundaryAllowsMarkupOverhead() {
        var plain = BookOnixTests.Project().ExportOnix(Options("<em>literal</em>", BookOnixCollateralTextFormat.PlainText), BookOnixTests.TestSchema());
        var note = XDocument.Load(new MemoryStream(plain.Bytes)).Descendants(Ns + "SalesRestrictionNote").Single();
        Assert.Equal("06", note.Attribute("textformat")!.Value); Assert.Empty(note.Elements()); Assert.Equal("<em>literal</em>", note.Value);
        var xhtml = BookOnixTests.Project().ExportOnix(Options("<p>" + string.Concat(Enumerable.Repeat("&amp;", 300)) + "</p>"), BookOnixTests.TestSchema());
        Assert.Equal(new string('&', 300), XDocument.Load(new MemoryStream(xhtml.Bytes)).Descendants(Ns + "SalesRestrictionNote").Single().Value);
    }

    [Theory]
    [InlineData("<script>alert(1)</script>")]
    [InlineData("<p onclick='bad()'>Text</p>")]
    [InlineData("<a href='file:///etc/passwd'>Text</a>")]
    [InlineData("<a href='javascript:bad()'>Text</a>")]
    [InlineData("<p xmlns='urn:other'>Text</p>")]
    [InlineData("<p>broken")]
    [InlineData("<!--comment--><p>Text</p>")]
    [InlineData("<p> </p>")]
    [InlineData("<!DOCTYPE p SYSTEM 'https://example.org/evil'><p>Text</p>")]
    public void UnsafeOrMalformedNotesAreRejectedWithoutChangingProject(string text) => Reject(text);

    [Fact]
    public void DecodedAndSourceLimitsAreIndependent() {
        Reject("<p>" + new string('x', 301) + "</p>");
        Reject("<p>" + string.Concat(Enumerable.Repeat("&#x1F600;", 151)) + "</p>");
        BookOnixTests.Project().ExportOnix(Options("<p>" + string.Concat(Enumerable.Repeat("&#x1F600;", 150)) + "</p>"), BookOnixTests.TestSchema());
        Reject("<p title='" + new string('x', 4096) + "'>Text</p>");
        var project = BookOnixTests.Project();
        Assert.Throws<ArgumentOutOfRangeException>(() => project.ExportOnix(Options("Text", (BookOnixCollateralTextFormat)99), BookOnixTests.TestSchema()));
    }

    private static void Reject(string text) {
        var project = BookOnixTests.Project(); var before = project.ToProjectBytes();
        Assert.ThrowsAny<ArgumentException>(() => project.ExportOnix(Options(text), BookOnixTests.TestSchema()));
        Assert.Equal(before, project.ToProjectBytes());
    }
}
