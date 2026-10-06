using OfficeIMO.Epub;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubChapterMergeMetadataContracts {
    private static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";

    [Theory]
    [InlineData(false, "#two", "#second-position")]
    [InlineData(true, "package.opf#two", "package.opf#second-position")]
    [InlineData(true, "#%74wo", "#second%2Dposition")]
    public void ExplicitRetargetingPreservesMetadataNodesAndTheirRefinements(bool firstHasId, string resourceTarget, string positionTarget) {
        var book = Prepared(firstHasId, resourceTarget, positionTarget);
        XElement[] metadata = book.GetPackageXml().Root!.Element(Opf + "metadata")!.Elements().Select(x => new XElement(x)).ToArray();
        book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions { RetargetPackageRefinements = true });
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        var package = reopened.GetPackageXml();
        XElement position = package.Root!.Element(Opf + "spine")!.Elements().Single();
        string retainedId = firstHasId ? "first-position" : "second-position";
        Assert.Equal(retainedId, (string?)position.Attribute("id"));
        Assert.Equal("one", (string?)position.Attribute("idref"));
        XElement[] actual = package.Root.Element(Opf + "metadata")!.Elements().ToArray();
        foreach (XElement original in metadata.Where(x => x.Attribute("id") != null)) {
            string id = (string)original.Attribute("id")!;
            XElement retained = actual.Single(x => (string?)x.Attribute("id") == id);
            if (id == "second-description" || id == "linked-description") original.SetAttributeValue("refines", "#one");
            if (id == "position-description") original.SetAttributeValue("refines", "#" + retainedId);
            Assert.True(XNode.DeepEquals(original, retained), id);
        }
        Assert.DoesNotContain(reopened.Manifest, item => item.Id == "two");
        Assert.Equal("second-start", reopened.Read().TableOfContents[1].Fragment);
    }

    [Theory]
    [InlineData("default")]
    [InlineData("extension")]
    [InlineData("style")]
    [InlineData("cancel")]
    public void RejectedMergeLeavesRefinementTargetsAndPositionIdsUnchanged(string failure) {
        var book = Prepared(false, "#two", "#second-position", failure == "extension");
        if (failure == "style") {
            XDocument content = book.GetContentXml("two");
            XNamespace html = "http://www.w3.org/1999/xhtml";
            content.Root!.Element(html + "head")!.Add(new XElement(html + "style", "p{color:red}"));
            book.SetContentXml("two", content);
        }
        byte[] before = book.Write().Bytes;
        using var cancellation = new CancellationTokenSource();
        if (failure == "cancel") cancellation.Cancel();
        Assert.ThrowsAny<Exception>(() => book.MergeChapters("one", "two", "second-start",
            new EpubChapterMergeOptions { RetargetPackageRefinements = failure != "default" }, cancellation.Token));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void RefinementRetargetingRequiresEpub3() {
        var book = EpubPublication.Create("Merge metadata", "en", version: EpubVersion.Epub2);
        book.AddChapter("one", "EPUB/one.xhtml", "One", "<p>First</p>");
        book.AddChapter("two", "EPUB/two.xhtml", "Two", "<p>Second</p>");
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.MergeChapters("one", "two", "second-start", new EpubChapterMergeOptions { RetargetPackageRefinements = true }));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubPublication Prepared(bool firstHasId, string resourceTarget, string positionTarget, bool extension = false) {
        var authored = EpubPublication.Create("Merge metadata", "en");
        authored.AddChapter("one", "EPUB/one.xhtml", "One", "<p>First</p>");
        authored.AddChapter("two", "EPUB/two.xhtml", "Two", "<p>Second</p>");
        var package = authored.GetPackageXml();
        var positions = package.Root!.Element(Opf + "spine")!.Elements().ToArray();
        if (firstHasId) positions[0].SetAttributeValue("id", "first-position");
        positions[1].SetAttributeValue("id", "second-position");
        package.Root.Element(Opf + "metadata")!.Add(
            new XElement(Opf + "meta", new XAttribute("id", "first-description"), new XAttribute("property", "dcterms:description"), new XAttribute("refines", "#one"), "First description"),
            new XElement(Opf + "meta", new XAttribute("id", "second-description"), new XAttribute("property", "dcterms:description"), new XAttribute("refines", resourceTarget), new XAttribute(XNamespace.Xml + "lang", "en"), "Second description"),
            new XElement(Opf + "meta", new XAttribute("id", "nested-description"), new XAttribute("property", "dcterms:description"), new XAttribute("refines", "#second-description"), "Annotation"),
            new XElement(Opf + "meta", new XAttribute("id", "position-description"), new XAttribute("property", "dcterms:description"), new XAttribute("refines", positionTarget), "Reading position"),
            new XElement(Opf + "link", new XAttribute("id", "linked-description"), new XAttribute("rel", "dcterms:description"), new XAttribute("href", "https://example.org/description.html"), new XAttribute("media-type", "text/html"), new XAttribute("refines", resourceTarget)));
        if (extension) package.Root.Element(Opf + "metadata")!.Add(new XElement(XName.Get("annotation", "urn:test:extension"), new XAttribute("refines", "#two"), "Keep scoped"));
        using var stream = new MemoryStream();
        stream.Write(authored.Write().Bytes);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, true)) {
            zip.GetEntry(authored.PackagePath)!.Delete();
            using var output = zip.CreateEntry(authored.PackagePath).Open();
            package.Save(output);
        }
        return EpubPublication.Load(new MemoryStream(stream.ToArray()));
    }
}
