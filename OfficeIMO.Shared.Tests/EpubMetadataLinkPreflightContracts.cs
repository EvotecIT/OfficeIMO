using OfficeIMO.Epub;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubMetadataLinkPreflightContracts {
    private static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";

    [Theory]
    [InlineData("record", "#c", "application/xml", "EPUB_PREFLIGHT_LINK_REFINES_FORBIDDEN")]
    [InlineData("alternate", "#c", "application/pdf", "EPUB_PREFLIGHT_LINK_REFINES_FORBIDDEN")]
    [InlineData("alternate dcterms:description", null, "text/html", "EPUB_PREFLIGHT_LINK_ALTERNATE_COMBINED")]
    [InlineData("voicing", null, "audio/mpeg", "EPUB_PREFLIGHT_LINK_REFINES_REQUIRED")]
    [InlineData("record", null, null, "EPUB_PREFLIGHT_LINK_MEDIA_TYPE_MISSING")]
    [InlineData("voicing", "#c", null, "EPUB_PREFLIGHT_LINK_MEDIA_TYPE_MISSING")]
    [InlineData("record", null, "application/xml", null)]
    [InlineData("alternate", null, "application/pdf", null)]
    [InlineData("voicing", "#c", "audio/mpeg", null)]
    [InlineData("dcterms:description", "#c", "text/html", null)]
    [InlineData("custom:record", "#c", null, null)]
    public void RetainedRelationshipsAreDiagnosedWithoutChangingBytes(string relation, string? refines, string? mediaType, string? error) {
        byte[] bytes = Package(relation, refines, mediaType, false);
        var book = EpubPublication.Load(new MemoryStream(bytes));
        var check = Assert.Single(book.Preflight().Checks, item => item.Code == "metadata-links");
        Assert.Equal(error == null ? EpubPreflightStatus.Passed : EpubPreflightStatus.Failed, check.Status);
        if (error != null) {
            var finding = Assert.Single(check.Diagnostics);
            Assert.Equal(error, finding.Code);
            Assert.Equal(book.PackagePath, finding.Path);
            Assert.Contains("linked-resource", finding.Message);
        }
        Assert.Equal(bytes, book.Write().Bytes);
    }

    [Fact]
    public void CollectionMetadataUsesTheSameRelationshipRules() {
        byte[] bytes = Package("record", "#c", "application/xml", true);
        var book = EpubPublication.Load(new MemoryStream(bytes));
        var check = Assert.Single(book.Preflight().Checks, item => item.Code == "metadata-links");
        Assert.Equal("EPUB_PREFLIGHT_LINK_REFINES_FORBIDDEN", Assert.Single(check.Diagnostics).Code);
        Assert.Equal(bytes, book.Write().Bytes);
    }

    private static byte[] Package(string relation, string? refines, string? mediaType, bool collection) {
        var book = EpubPublication.Create("Metadata links", "en");
        book.AddChapter("c", "EPUB/c.xhtml", "Chapter", "<h1>Chapter</h1>");
        XDocument package = book.GetPackageXml();
        package.Root!.SetAttributeValue("prefix", "custom: https://example.org/vocabulary/");
        var link = new XElement(Opf + "link", new XAttribute("id", "linked-resource"),
            new XAttribute("rel", relation), new XAttribute("href", "https://example.org/resource"));
        link.SetAttributeValue("refines", refines);
        link.SetAttributeValue("media-type", mediaType);
        if (collection) package.Root.Add(new XElement(Opf + "collection", new XAttribute("role", "https://example.org/collection"),
            new XElement(Opf + "metadata", link), new XElement(Opf + "link", new XAttribute("href", "c.xhtml"))));
        else package.Root.Element(Opf + "metadata")!.Add(link);
        return EpubWritingContracts.ReplaceEntry(book.Write().Bytes, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()));
    }
}
