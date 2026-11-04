using System.Threading;
using OfficeIMO.Epub;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubIndexAuthoringContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private static readonly XNamespace Ops = "http://www.idpf.org/2007/ops";

    [Fact]
    public void LocatorAliasesToOneResourceShareTheirValidatedTarget() {
        using var output = new MemoryStream();
        byte[] original = Book().Write().Bytes;
        output.Write(original, 0, original.Length);
        using (var archive = new ZipArchive(output, ZipArchiveMode.Update, true)) {
            ZipArchiveEntry packageEntry = archive.GetEntry("EPUB/package.opf")!;
            XDocument package;
            using (var source = packageEntry.Open()) package = XDocument.Load(source);
            XNamespace opf = package.Root!.Name.Namespace;
            XElement manifest = package.Root.Element(opf + "manifest")!;
            XElement spine = package.Root.Element(opf + "spine")!;
            XElement declared = manifest.Elements(opf + "item").Single(item => (string?)item.Attribute("id") == "source");
            foreach (int index in Enumerable.Range(0, 12)) {
                string id = "alias-" + index;
                XElement alias = new XElement(declared); alias.SetAttributeValue("id", id);
                manifest.Add(alias);
                spine.Add(new XElement(opf + "itemref", new XAttribute("idref", id)));
            }
            packageEntry.Delete();
            using var destination = archive.CreateEntry("EPUB/package.opf", CompressionLevel.Optimal).Open();
            package.Save(destination);
        }
        var book = EpubPublication.Load(new MemoryStream(output.ToArray()));
        EpubIndexLocator[] locators = Enumerable.Range(0, 12).Select(index => new EpubIndexLocator {
            ManifestId = "alias-" + index, FragmentId = "topic", Label = "Page " + index
        }).ToArray();

        book.AddIndexEntry("index", "entries", "term", "Reading", locators);

        Assert.Equal(12, book.GetContentXml("index").Descendants(Html + "a").Count());
        Assert.Equal(book.GetResourceBytes("source"), book.GetResourceBytes("alias-0"));
    }

    [Fact]
    public void NestedIndexPreservesPublisherOrderAndResolvesBodyAndDocumentTargets() {
        var book = Book();
        byte[] source = book.GetResourceBytes("source");
        var content = book.GetContentXml("index");
        content.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("href", "../assets/")));
        book.SetContentXml("index", content);
        book.AddIndexEntry("index", "entries", "term", "Publishing & books", Array.Empty<EpubIndexLocator>(), "subentries");
        book.AddIndexEntry("index", "subentries", "child", "Reading order", new[] {
            new EpubIndexLocator { ManifestId = "source", FragmentId = "topic", Label = "12" },
            new EpubIndexLocator { ManifestId = "source", Label = "Whole chapter" }
        });
        book.AddIndexEntry("index", "entries", "cross", "See publishing", new[] {
            new EpubIndexLocator { ManifestId = "index", FragmentId = "term", Label = "Publishing & books" }
        });
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal(source, reopened.GetResourceBytes("source"));
        var index = reopened.GetContentXml("index");
        var section = index.Descendants(Html + "section").Single();
        Assert.Equal("doc-index", (string?)section.Attribute("role"));
        Assert.Equal("index", (string?)section.Attribute(Ops + "type"));
        Assert.Equal("Publishing & books", ById(index, "term").Element(Html + "span")!.Value);
        Assert.Equal("term", (string?)ById(index, "child").Parent!.Parent!.Attribute("id"));
        var links = ById(index, "child").Elements(Html + "a").ToArray();
        Assert.Equal(new[] { "12", "Whole chapter" }, links.Select(e => e.Value));
        var target = EpubReference.Resolve("EPUB/back/index.xhtml", "../assets/", links[0].Attribute("href")!.Value);
        Assert.Equal("EPUB/text/source.xhtml", target.ContainerPath); Assert.Equal("topic", target.Fragment);
        Assert.Null(EpubReference.Resolve("EPUB/back/index.xhtml", "../assets/", links[1].Attribute("href")!.Value).Fragment);
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("head")]
    [InlineData("duplicate")]
    [InlineData("non-spine")]
    public void InvalidLocatorOrIdentifierLeavesIndexUnchanged(string failure) {
        var book = Book();
        byte[] before = book.Write().Bytes;
        var locator = new EpubIndexLocator { ManifestId = failure == "non-spine" ? "nav" : "source", FragmentId = failure == "missing" ? "absent" : failure == "head" ? "metadata" : "topic", Label = "12" };
        Assert.Throws<InvalidDataException>(() => book.AddIndexEntry("index", "entries", failure == "duplicate" ? "entries" : "term", "Term", new[] { locator }));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void CancellationAndLocatorCountAreBoundedBeforeMutation() {
        var book = Book(); byte[] before = book.Write().Bytes;
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => book.AddIndexEntry("index", "entries", "term", "Term", Array.Empty<EpubIndexLocator>(), "children", cancellation.Token));
        Assert.Throws<ArgumentOutOfRangeException>(() => book.AddIndexEntry("index", "entries", "term", "Term", new EpubIndexLocator[1025]));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static XElement ById(XDocument document, string id) => document.Descendants().Single(e => (string?)e.Attribute("id") == id);
    private static EpubPublication Book() {
        var book = EpubPublication.Create("Index", "en");
        book.AddChapter("source", "EPUB/text/source.xhtml", "Chapter", "<h1 id='topic'>Reading order</h1><p>Text.</p>");
        book.AddChapter("index", "EPUB/back/index.xhtml", "Index", "<section><h1>Index</h1><ul id='entries'/></section>");
        var source = book.GetContentXml("source");
        source.Root!.Element(Html + "head")!.Element(Html + "title")!.SetAttributeValue("id", "metadata");
        book.SetContentXml("source", source);
        return book;
    }
}
