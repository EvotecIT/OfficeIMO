using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Epub;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubWriterRetentionContracts {
    private static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Theory]
    [InlineData("add")]
    [InlineData("xml")]
    [InlineData("update")]
    public void ContentSafety_RejectsRootSvgEventHandlers(string operation) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        byte[] safe = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg'><text>Safe</text></svg>");
        byte[] scripted = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' onload='alert(1)'><text>Unsafe</text></svg>");
        if (operation == "add") {
            book.AddResource("svg", "EPUB/page.svg", "image/svg+xml", scripted);
            Assert.Throws<NotSupportedException>(() => book.Write());
        } else {
            book.AddResource("svg", "EPUB/page.svg", "image/svg+xml", safe);
            if (operation == "xml") Assert.Throws<NotSupportedException>(() => book.SetContentXml("svg", XDocument.Parse(Encoding.UTF8.GetString(scripted))));
            else { book.UpdateResource("svg", scripted); Assert.Throws<NotSupportedException>(() => book.Write()); }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Navigation_AcceptsFragmentIdsOnContentRoots(bool svg) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        string path = svg ? "EPUB/page.svg" : "EPUB/first.xhtml";
        if (svg) {
            book.AddResource("svg", path, "image/svg+xml", Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' id='root'><text>Root</text></svg>"));
            book.AddSpineItem("svg");
        } else {
            XDocument content = book.GetContentXml("first");
            content.Root!.SetAttributeValue("id", "root");
            book.SetContentXml("first", content);
        }
        book.SetNavigation(new[] { new EpubNavigationEntry("Root", path + "#root") });
        Assert.NotEmpty(book.Write().Bytes);
    }

    [Fact]
    public void OrdinaryNavigationHyperlinks_RequireLocalForeignContentInTheSpine() {
        EpubPublication book = EpubWritingContracts.CreateBook();
        book.AddResource("image", "EPUB/image.png", "image/png", new byte[] { 1 });
        XDocument nav = book.GetContentXml("navigation");
        nav.Root!.Element(Html + "body")!.Add(new XElement(Html + "p", new XElement(Html + "a", new XAttribute("href", "image.png"), "Image")));
        book.SetContentXml("navigation", nav);
        Assert.Throws<InvalidDataException>(() => book.Write());
        book.Manifest.Single(item => item.Id == "image").FallbackId = "first";
        book.AddSpineItem("image", linear: false);
        Assert.NotEmpty(book.Write().Bytes);
    }

    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void Create_CountsMandatoryEntriesAgainstRetention(EpubVersion version) {
        Assert.Throws<InvalidDataException>(() => EpubPublication.Create("Small", version: version,
            retentionLimits: new EpubPublicationLoadOptions { MaxEntries = 3 }));
        EpubPublication book = EpubPublication.Create("Small", version: version, retentionLimits: new EpubPublicationLoadOptions { MaxEntries = 4 });
        Assert.Throws<InvalidDataException>(() => book.AddChapter("c", "EPUB/chapter.xhtml", "Chapter", "<p>Body</p>"));
        Assert.Empty(book.Spine);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LoadedRetention_RejectsNewEntriesBeforeMutatingThePublication(bool chapter) {
        byte[] source = EpubWritingContracts.CreateBook().Write().Bytes;
        using var archive = new ZipArchive(new MemoryStream(source));
        EpubPublication book = EpubPublication.Load(new MemoryStream(source), new EpubPublicationLoadOptions { MaxEntries = archive.Entries.Count });
        Action add = chapter ? () => book.AddChapter("extra", "EPUB/extra.xhtml", "Extra", "<p>Extra</p>") :
            () => book.AddResource("extra", "EPUB/extra.css", "text/css", Encoding.UTF8.GetBytes("body{}"));
        Assert.Throws<InvalidDataException>(add);
        Assert.Equal(source, book.Write().Bytes);
        Assert.DoesNotContain(book.Manifest, item => item.Id == "extra");
        book.UpdateResource("style", Encoding.UTF8.GetBytes("body{color:blue}"));
        Assert.NotEmpty(book.Write().Bytes);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Epub2_RejectsNestedFlatNavigationSectionsWithoutDroppingChildren(bool landmarks) {
        EpubPublication book = EpubWritingContracts.CreateBook(EpubVersion.Epub2);
        byte[] before = book.Write().Bytes;
        var nested = new[] { new EpubNavigationEntry("First", "EPUB/first.xhtml", new[] {
            new EpubNavigationEntry("Second", "EPUB/second.xhtml", semanticType: "text") }, semanticType: "text") };
        Assert.Throws<NotSupportedException>(() => book.SetNavigation(new[] { new EpubNavigationEntry("Contents", "EPUB/first.xhtml") },
            pageList: landmarks ? null : nested, landmarks: landmarks ? nested : null));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Theory]
    [InlineData("sections/", false)]
    [InlineData("sections/", true)]
    [InlineData("sections/base.xhtml?language=en", false)]
    [InlineData("../other%20dir/", true)]
    [InlineData(".", false)]
    [InlineData(".", true)]
    [InlineData("sections/..", false)]
    [InlineData("sections/..", true)]
    public void NavigationEdits_ResolveGeneratedLinksAgainstRetainedHtmlBase(string baseHref, bool append) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        XDocument nav = book.GetContentXml("navigation");
        nav.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("href", baseHref)));
        var documentUri = new Uri("epub://package/EPUB/nav.xhtml");
        var baseUri = new Uri(documentUri, baseHref);
        foreach (XElement anchor in nav.Descendants(Html + "a")) anchor.SetAttributeValue("href",
            baseUri.MakeRelativeUri(new Uri(documentUri, anchor.Attribute("href")!.Value)).OriginalString);
        book.SetContentXml("navigation", nav);
        if (append) book.AddChapter("third", "EPUB/third.xhtml", "Third", "<p>Third</p>");
        else book.SetNavigation(new[] { new EpubNavigationEntry("First", "EPUB/first.xhtml#root") });
        nav = book.GetContentXml("navigation");
        Assert.Equal(baseHref, nav.Root!.Element(Html + "head")!.Element(Html + "base")!.Attribute("href")!.Value);
        string expected = append ? "EPUB/third.xhtml" : "EPUB/first.xhtml";
        Assert.Equal(expected, EpubReference.Resolve("EPUB/nav.xhtml", baseHref, nav.Descendants(Html + "a").Last().Attribute("href")!.Value).ContainerPath);
        Assert.Equal(expected, Uri.UnescapeDataString(new Uri(baseUri, nav.Descendants(Html + "a").Last().Attribute("href")!.Value).AbsolutePath.TrimStart('/')));
        if (!append) {
            XDocument content = book.GetContentXml("first"); content.Root!.SetAttributeValue("id", "root"); book.SetContentXml("first", content);
        }
        Assert.NotEmpty(book.Write().Bytes);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NavigationEdits_RejectExternalBasesBeforeChangingThePublication(bool append) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        XDocument nav = book.GetContentXml("navigation");
        nav.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("href", "https://example.org/")));
        book.SetContentXml("navigation", nav);
        byte[] before = book.GetResourceBytes("navigation");
        int count = book.Manifest.Count;
        Assert.Throws<NotSupportedException>(() => {
            if (append) book.AddChapter("third", "EPUB/third.xhtml", "Third", "<p>Third</p>");
            else book.SetNavigation(new[] { new EpubNavigationEntry("First", "EPUB/first.xhtml") });
        });
        Assert.Equal(before, book.GetResourceBytes("navigation"));
        Assert.Equal(count, book.Manifest.Count);
    }

    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void Write_HonorsStoredCompressionPolicyOnUneditedImports(EpubVersion version) {
        byte[] source = EpubWritingContracts.CreateBook(version).Write().Bytes;
        EpubPublication book = EpubPublication.Load(new MemoryStream(source));
        EpubWriteResult stored = book.Write(new EpubWriteOptions { CompressEntries = false });
        Assert.False(stored.Report.UsedOriginalPackage);
        using var archive = new ZipArchive(new MemoryStream(stored.Bytes));
        Assert.All(archive.Entries, entry => Assert.Equal(entry.Length, entry.CompressedLength));
        Assert.Equal(ReadEntry(source, book.PackagePath), ReadEntry(stored.Bytes, book.PackagePath));
        Assert.Equal(source, book.Write().Bytes);
    }

    [Fact]
    public void MetadataEdits_RejectAmbiguousModificationStampsBeforeSaving() {
        EpubPublication book = EpubWritingContracts.CreateBook();
        byte[] source = book.Write().Bytes;
        XDocument package = XDocument.Parse(Encoding.UTF8.GetString(ReadEntry(source, book.PackagePath)));
        XElement stamp = package.Descendants(Opf + "meta").Single(meta => (string?)meta.Attribute("property") == "dcterms:modified");
        stamp.AddAfterSelf(new XElement(stamp));
        book = EpubPublication.Load(new MemoryStream(EpubWritingContracts.ReplaceEntry(source, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()))));
        book.Title = "Edited";
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, true);
        Assert.Throws<InvalidDataException>(() => book.Save(destination));
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
    }

    [Fact]
    public void MetadataEdits_PreserveModificationStampIdentityAndExtensions() {
        EpubPublication book = EpubWritingContracts.CreateBook();
        byte[] source = book.Write().Bytes;
        XDocument package = XDocument.Parse(Encoding.UTF8.GetString(ReadEntry(source, book.PackagePath)));
        XElement stamp = package.Descendants(Opf + "meta").Single(meta => (string?)meta.Attribute("property") == "dcterms:modified");
        stamp.SetAttributeValue("id", "modified"); stamp.SetAttributeValue(XName.Get("flag", "urn:retained"), "keep");
        package.Root!.Element(Opf + "metadata")!.Add(new XElement(Opf + "meta", new XAttribute("property", "schema:description"), new XAttribute("refines", "#modified"), "Revision date"));
        book = EpubPublication.Load(new MemoryStream(EpubWritingContracts.ReplaceEntry(source, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()))));
        book.Title = "Edited";
        byte[] output = book.Write(new EpubWriteOptions { ModifiedAt = DateTimeOffset.Parse("2020-01-02T03:04:05Z") }).Bytes;
        package = XDocument.Parse(Encoding.UTF8.GetString(ReadEntry(output, book.PackagePath)));
        stamp = package.Descendants(Opf + "meta").Single(meta => (string?)meta.Attribute("property") == "dcterms:modified");
        Assert.Equal("modified", stamp.Attribute("id")!.Value);
        Assert.Equal("keep", stamp.Attribute(XName.Get("flag", "urn:retained"))!.Value);
        Assert.Equal("2020-01-02T03:04:05Z", stamp.Value);
        Assert.Contains(package.Descendants(Opf + "meta"), meta => (string?)meta.Attribute("refines") == "#modified");
    }

    private static byte[] ReadEntry(byte[] data, string path) {
        using var archive = new ZipArchive(new MemoryStream(data)); using Stream input = archive.GetEntry(path)!.Open();
        using var output = new MemoryStream(); input.CopyTo(output); return output.ToArray();
    }

    [Theory]
    [InlineData(EpubVersion.Epub2, false)]
    [InlineData(EpubVersion.Epub2, true)]
    [InlineData(EpubVersion.Epub3, false)]
    [InlineData(EpubVersion.Epub3, true)]
    public void MixedCaseMediaTypes_PreserveUneditedImportsAndSupportContentEdits(EpubVersion version, bool allResources) {
        EpubPublication book = EpubWritingContracts.CreateBook(version);
        byte[] source = book.Write().Bytes;
        XDocument package = book.GetPackageXml();
        foreach (XElement item in package.Descendants(Opf + "item").Where(item => allResources || (string?)item.Attribute("id") == "navigation"))
            item.Attribute("media-type")!.Value = item.Attribute("media-type")!.Value.ToUpperInvariant();
        source = EpubWritingContracts.ReplaceEntry(source, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()));
        book = EpubPublication.Load(new MemoryStream(source));
        Assert.Equal(source, book.Write().Bytes);
        string declared = book.Manifest.Single(item => item.Id == "navigation").MediaType;
        book.Title = "Edited";
        XDocument content = book.GetContentXml("first");
        content.Root!.Element(Html + "body")!.Add(new XElement(Html + "p", "Edit"));
        book.SetContentXml("first", content);
        book.AddChapter("third", "EPUB/third.xhtml", "Third", "<p>Third</p>", new[] { "style" });
        book.AddResource("cover", "EPUB/cover.png", "IMAGE/PNG", new byte[] { 1 });
        book.SetCoverImage("cover");
        EpubPublication loaded = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal(declared, loaded.Manifest.Single(item => item.Id == "navigation").MediaType);
        Assert.Equal("IMAGE/PNG", loaded.Manifest.Single(item => item.Id == "cover").MediaType);
    }

    [Fact]
    public void Epub2_IdentitySynchronizationEnforcesRetainedNcxEntryBytes() {
        EpubPublication book = EpubWritingContracts.CreateBook(EpubVersion.Epub2);
        XDocument ncx = book.GetContentXml("navigation");
        ncx.Root!.AddFirst(new XComment(new string('x', 4096)));
        book.UpdateResource("navigation", Encoding.UTF8.GetBytes(ncx.ToString()));
        byte[] source = book.Write().Bytes;
        using var archive = new ZipArchive(new MemoryStream(source));
        book = EpubPublication.Load(new MemoryStream(source), new EpubPublicationLoadOptions { MaxEntryBytes = archive.Entries.Max(entry => entry.Length) });
        book.Identifier = "urn:example:" + new string('a', 1500);
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, true);
        Assert.Throws<InvalidDataException>(() => book.Save(destination));
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
    }

    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void NavigationDeclarations_RejectWrongMediaTypesOnSaveAndAuthoring(EpubVersion version) {
        EpubPublication book = EpubWritingContracts.CreateBook(version);
        book.Manifest.Single(item => item.Id == "navigation").MediaType = "text/plain";
        Assert.Throws<InvalidDataException>(() => book.Write());
        int count = book.Manifest.Count;
        Assert.Throws<InvalidDataException>(() => book.AddChapter("third", "EPUB/third.xhtml", "Third", "<p>Third</p>"));
        Assert.Equal(count, book.Manifest.Count);
    }

    [Fact]
    public void Epub2_RejectsForeignNavigationRootsEvenWhenNcxChildrenRemain() {
        EpubPublication book = EpubWritingContracts.CreateBook(EpubVersion.Epub2);
        XDocument ncx = book.GetContentXml("navigation");
        ncx.Root!.Name = XName.Get("ncx", "urn:foreign");
        ncx.Root.Attribute("xmlns")?.Remove();
        book.UpdateResource("navigation", Encoding.UTF8.GetBytes(ncx.ToString()));
        Assert.Throws<InvalidDataException>(() => book.Write());
    }

    [Fact]
    public void Epub2_PageListClearingRejectsUnavoidableHeaderAndExtensionLoss() {
        EpubPublication book = EpubWritingContracts.CreateBook(EpubVersion.Epub2);
        book.SetNavigation(new[] { new EpubNavigationEntry("First", "EPUB/first.xhtml") },
            pageList: new[] { new EpubNavigationEntry("2", "EPUB/second.xhtml") });
        XDocument ncx = book.GetContentXml("navigation");
        XNamespace ns = "http://www.daisy.org/z3986/2005/ncx/";
        ncx.Root!.Element(ns + "pageList")!.Add(new XElement(XName.Get("retained", "urn:extension"), "Keep me"));
        book.UpdateResource("navigation", Encoding.UTF8.GetBytes(ncx.ToString()));
        byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.SetNavigation(new[] { new EpubNavigationEntry("First", "EPUB/first.xhtml") }, pageList: Array.Empty<EpubNavigationEntry>()));
        Assert.Equal(before, book.Write().Bytes);
    }
}
