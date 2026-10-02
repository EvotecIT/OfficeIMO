using System.Text;
using System.Xml.Linq;
using OfficeIMO.Epub;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubWriterFeedbackContracts {
    private static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Fact]
    public void CoverSelection_OmitsEmptyPropertyListsAndPreservesOtherProperties() {
        EpubPublication book = EpubWritingContracts.CreateBook();
        book.AddResource("cover", "EPUB/cover.png", "image/png", new byte[] { 1 });
        book.SetCoverImage("cover");
        EpubPublication loaded = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.DoesNotContain(loaded.GetPackageXml().Descendants().Attributes("properties"), attribute => string.IsNullOrWhiteSpace(attribute.Value));
        Assert.Equal("cover-image", loaded.Manifest.Single(item => item.Id == "cover").Properties);
        Assert.Equal("nav", loaded.Manifest.Single(item => item.Id == "navigation").Properties);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImportedResource_RejectsNewScriptDeclarationsWithoutPayloadChanges(bool property) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        book.AddResource("payload", "EPUB/payload.js", "text/plain", Encoding.UTF8.GetBytes("alert(1)"));
        book = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        EpubManifestItem chapter = book.Manifest.Single(item => item.Id == (property ? "first" : "payload"));
        if (property) chapter.Properties = "scripted"; else chapter.MediaType = "application/javascript";
        Assert.Throws<NotSupportedException>(() => book.Write());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Removal_RejectsRetainedPackageReferences(bool refinement) {
        EpubPublication book = EpubWritingContracts.CreateBook(refinement ? EpubVersion.Epub3 : EpubVersion.Epub2);
        byte[] source = book.Write().Bytes;
        XDocument package = book.GetPackageXml();
        if (refinement) package.Root!.Element(Opf + "metadata")!.Add(new XElement(Opf + "meta",
            new XAttribute("property", "media:duration"), new XAttribute("refines", "#style"), "0:01:00"));
        else package.Descendants(Opf + "item").Single(item => (string?)item.Attribute("id") == "first").SetAttributeValue("fallback-style", "style");
        book = EpubPublication.Load(new MemoryStream(EpubWritingContracts.ReplaceEntry(source, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()))));
        Assert.Throws<InvalidOperationException>(() => book.RemoveResource("style"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PackageResourceUrls_ProtectRetainedMetadataPayloads(bool refinement) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        book.AddResource("record", "EPUB/record.xml", "application/xml", Encoding.UTF8.GetBytes("<record/>"));
        byte[] source = book.Write().Bytes;
        XDocument package = book.GetPackageXml();
        package.Root!.Element(Opf + "metadata")!.Add(refinement ? new XElement(Opf + "meta", new XAttribute("property", "schema:alternateName"),
            new XAttribute("refines", "record.xml"), "Record") : new XElement(Opf + "link", new XAttribute("rel", "record"), new XAttribute("href", "record.xml")));
        book = EpubPublication.Load(new MemoryStream(EpubWritingContracts.ReplaceEntry(source, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()))));
        Assert.Throws<InvalidOperationException>(() => book.RemoveResource("record"));
        book.Title = "Edited";
        Assert.Contains("EPUB/record.xml", EpubPublication.Load(new MemoryStream(book.Write().Bytes)).EntryPaths);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Write_RejectsDanglingRetainedPackageReferences(bool refinement) {
        EpubPublication book = EpubWritingContracts.CreateBook(refinement ? EpubVersion.Epub3 : EpubVersion.Epub2);
        byte[] source = book.Write().Bytes;
        XDocument package = book.GetPackageXml();
        if (refinement) package.Root!.Element(Opf + "metadata")!.Add(new XElement(Opf + "meta",
            new XAttribute("property", "media:duration"), new XAttribute("refines", "#missing"), "0:01:00"));
        else package.Descendants(Opf + "item").Single(item => (string?)item.Attribute("id") == "first").SetAttributeValue("fallback-style", "missing");
        book = EpubPublication.Load(new MemoryStream(EpubWritingContracts.ReplaceEntry(source, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()))));
        book.Title = "Edited";
        Assert.Throws<InvalidDataException>(() => book.Write());
    }

    [Fact]
    public void ContentEdit_UsesTheEffectiveHtmlBaseForRemoteResourceDeclarations() {
        EpubPublication book = EpubWritingContracts.CreateBook();
        XDocument content = book.GetContentXml("first");
        content.Root!.Element(Html + "head")!.AddFirst(new XElement(Html + "base", new XAttribute("href", "https://example.org/images/")));
        content.Root.Element(Html + "body")!.Add(new XElement(Html + "img", new XAttribute("src", "image.png"), new XAttribute("alt", "Remote image")));
        book.SetContentXml("first", content);
        EpubPublication loaded = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Contains("remote-resources", loaded.Manifest.Single(item => item.Id == "first").Properties!.Split(' '));
    }

    [Fact]
    public void ClearingLandmarks_RetainsUnknownGuideXml() {
        EpubPublication book = EpubWritingContracts.CreateBook(EpubVersion.Epub2);
        byte[] source = book.Write().Bytes;
        XDocument package = book.GetPackageXml();
        XNamespace extension = "urn:retained";
        package.Root!.Add(new XElement(Opf + "guide", new XAttribute(extension + "flag", "keep"),
            new XElement(extension + "data", "opaque"), new XElement(Opf + "reference", new XAttribute("type", "text"), new XAttribute("href", "first.xhtml"))));
        book = EpubPublication.Load(new MemoryStream(EpubWritingContracts.ReplaceEntry(source, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()))));
        book.SetNavigation(new[] { new EpubNavigationEntry("Start", "EPUB/first.xhtml") }, landmarks: Array.Empty<EpubNavigationEntry>());
        EpubPublication loaded = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        XElement guide = loaded.GetPackageXml().Root!.Element(Opf + "guide")!;
        Assert.NotNull(guide);
        Assert.Equal("keep", guide.Attribute(extension + "flag")!.Value);
        Assert.Equal("opaque", guide.Element(extension + "data")!.Value);
        Assert.Empty(guide.Elements(Opf + "reference"));
    }

    [Theory]
    [InlineData(EpubVersion.Epub2, false)]
    [InlineData(EpubVersion.Epub2, true)]
    [InlineData(EpubVersion.Epub3, false)]
    [InlineData(EpubVersion.Epub3, true)]
    public void Navigation_RejectsTargetsOutsideTheSpine(EpubVersion version, bool unmanifested) {
        EpubPublication book = EpubWritingContracts.CreateBook(version);
        book.AddResource("outside", "EPUB/outside.xhtml", "application/xhtml+xml", Encoding.UTF8.GetBytes(EpubIntegrityFixtures.Xhtml("<p>Outside</p>")));
        byte[] source = book.Write().Bytes;
        if (unmanifested) {
            XDocument package = book.GetPackageXml();
            package.Descendants(Opf + "item").Single(item => (string?)item.Attribute("id") == "outside").Remove();
            source = EpubWritingContracts.ReplaceEntry(source, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()));
        }
        book = EpubPublication.Load(new MemoryStream(source));
        Assert.Throws<InvalidDataException>(() => book.SetNavigation(new[] { new EpubNavigationEntry("Outside", "EPUB/outside.xhtml") }));
    }

    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void Navigation_RejectsEmptyTocWithoutChangingThePublication(EpubVersion version) {
        EpubPublication book = EpubWritingContracts.CreateBook(version);
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.SetNavigation(Array.Empty<EpubNavigationEntry>()));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void Landmarks_RequireSemanticTypesOnNestedLinks() {
        EpubPublication book = EpubWritingContracts.CreateBook();
        Assert.Throws<InvalidDataException>(() => book.SetNavigation(new[] { new EpubNavigationEntry("First", "EPUB/first.xhtml") },
            landmarks: new[] { new EpubNavigationEntry("Start", "EPUB/first.xhtml", new[] { new EpubNavigationEntry("Nested", "EPUB/second.xhtml") }, "bodymatter") }));
    }

    [Theory]
    [InlineData("resource")]
    [InlineData("spine")]
    [InlineData("manifest")]
    [InlineData("overlay")]
    [InlineData("add-spine")]
    public void Epub2_RejectsEpub3OnlyDeclarations(string operation) {
        EpubPublication book = EpubWritingContracts.CreateBook(EpubVersion.Epub2);
        Action edit = operation == "resource" ? () => book.AddResource("extra", "EPUB/extra.css", "text/css", new byte[] { 1 }, "remote-resources") :
            operation == "add-spine" ? () => book.AddSpineItem("style", properties: "page-spread-left") :
            operation == "spine" ? () => book.Spine[0].Properties = "page-spread-left" :
            operation == "overlay" ? () => book.Manifest[0].MediaOverlayId = "first" :
            () => book.Manifest[0].Properties = "nav";
        Assert.Throws<NotSupportedException>(edit);
    }

    [Theory]
    [InlineData("title")]
    [InlineData("metadata")]
    [InlineData("manifest")]
    public void PackageEdits_EnforceMetadataRetentionWithoutCommittingFailedChanges(string operation) {
        EpubPublication book = EpubPublication.Create("Small", retentionLimits: new EpubPublicationLoadOptions { MaxMetadataBytes = 2048 });
        book.AddChapter("chapter", "EPUB/chapter.xhtml", "Chapter", "<p>Body</p>");
        XDocument before = book.GetPackageXml();
        string large = new string('x', 4096);
        Action edit = operation == "title" ? () => book.Title = large : operation == "metadata" ?
            () => book.AddDublinCoreMetadata("description", large) : () => book.Manifest[0].Properties = large;
        Assert.Throws<InvalidDataException>(edit);
        Assert.True(XNode.DeepEquals(before, book.GetPackageXml()));
    }

    [Fact]
    public void PackageEdits_CountLiveXmlAgainstExpandedRetention() {
        EpubPublication source = EpubWritingContracts.CreateBook();
        byte[] bytes = source.Write().Bytes;
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
        long expanded = archive.Entries.Sum(entry => entry.Length);
        EpubPublication book = EpubPublication.Load(new MemoryStream(bytes), new EpubPublicationLoadOptions { MaxExpandedBytes = expanded + 128 });
        XDocument before = book.GetPackageXml();
        Assert.Throws<InvalidDataException>(() => book.AddDublinCoreMetadata("description", new string('x', 4096)));
        Assert.True(XNode.DeepEquals(before, book.GetPackageXml()));
    }

    [Fact]
    public void FailedChapterEdit_DoesNotRetainPayloadOrPartialNavigation() {
        EpubPublication book = EpubPublication.Create("Small", retentionLimits: new EpubPublicationLoadOptions { MaxMetadataBytes = 2048 });
        book.AddChapter("first", "EPUB/first.xhtml", "First", "<p>Original</p>");
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.AddChapter(new string('x', 4096), "EPUB/second.xhtml", "Second", "<p>New</p>"));
        Assert.Equal(before, book.Write().Bytes);
        Assert.DoesNotContain("EPUB/second.xhtml", book.EntryPaths);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImportedScriptDeclarations_ArePreservedWhenUnchanged(bool property) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        book.AddResource("payload", "EPUB/payload.js", "text/plain", Encoding.UTF8.GetBytes("alert(1)"));
        byte[] source = book.Write().Bytes;
        XDocument package = book.GetPackageXml();
        XElement item = package.Descendants(Opf + "item").Single(element => (string?)element.Attribute("id") == (property ? "first" : "payload"));
        item.SetAttributeValue(property ? "properties" : "media-type", property ? "scripted" : "application/javascript");
        book = EpubPublication.Load(new MemoryStream(EpubWritingContracts.ReplaceEntry(source, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()))));
        book.Title = "Metadata edit";
        EpubPublication output = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal(property ? "scripted" : "application/javascript", property ? output.Manifest.Single(resource => resource.Id == "first").Properties :
            output.Manifest.Single(resource => resource.Id == "payload").MediaType);
        if (property) output.Manifest.Single(resource => resource.Id == "first").Properties = null;
        else output.Manifest.Single(resource => resource.Id == "payload").MediaType = "text/plain";
        Assert.Throws<NotSupportedException>(() => output.Write());
    }

    [Theory]
    [InlineData("empty-toc")]
    [InlineData("landmark")]
    [InlineData("outside-spine")]
    [InlineData("external-toc")]
    public void Write_PreflightsNavigationChangedThroughResourceAndSpineApis(string change) {
        EpubPublication book = EpubWritingContracts.CreateBook();
        if (change == "outside-spine") book.RemoveSpineItem(1);
        else {
            XDocument navigation = book.GetContentXml("navigation");
            XElement body = navigation.Root!.Element(Html + "body")!;
            if (change == "empty-toc") body.Descendants(Html + "ol").First().RemoveNodes();
            else if (change == "external-toc") body.Descendants(Html + "a").First().SetAttributeValue("href", "https://example.org/chapter.xhtml");
            else body.Add(new XElement(Html + "nav", new XAttribute(XName.Get("type", "http://www.idpf.org/2007/ops"), "landmarks"),
                new XElement(Html + "ol", new XElement(Html + "li", new XElement(Html + "a", new XAttribute("href", "first.xhtml"), "First")))));
            book.SetContentXml("navigation", navigation);
        }
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, true);
        Assert.Throws<InvalidDataException>(() => book.Save(destination));
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
    }

    [Theory]
    [InlineData("toc")]
    [InlineData("page-list")]
    [InlineData("guide")]
    [InlineData("tour")]
    public void Epub2_PreflightsRetainedNavigationTargetsAfterSpineEdits(string mechanism) {
        EpubPublication book = EpubWritingContracts.CreateBook(EpubVersion.Epub2);
        if (mechanism != "toc") book.SetNavigation(new[] { new EpubNavigationEntry("First", "EPUB/first.xhtml") },
            pageList: mechanism == "page-list" ? new[] { new EpubNavigationEntry("2", "EPUB/second.xhtml") } : null,
            landmarks: mechanism == "guide" ? new[] { new EpubNavigationEntry("Start", "EPUB/second.xhtml", semanticType: "text") } : null);
        if (mechanism == "tour") {
            byte[] source = book.Write().Bytes;
            XDocument package = book.GetPackageXml();
            package.Root!.Add(new XElement(Opf + "tours", new XElement(Opf + "tour", new XAttribute("id", "tour"), new XAttribute("title", "Tour"),
                new XElement(Opf + "site", new XAttribute("title", "Second"), new XAttribute("href", "second.xhtml")))));
            book = EpubPublication.Load(new MemoryStream(EpubWritingContracts.ReplaceEntry(source, book.PackagePath, Encoding.UTF8.GetBytes(package.ToString()))));
        }
        book.RemoveSpineItem(1);
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 }, true);
        Assert.Throws<InvalidDataException>(() => book.Save(destination));
        Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
    }

    [Theory]
    [InlineData(EpubVersion.Epub2, false)]
    [InlineData(EpubVersion.Epub2, true)]
    [InlineData(EpubVersion.Epub3, false)]
    [InlineData(EpubVersion.Epub3, true)]
    public void UneditedImports_UseOriginalByteLengthsAtExactRetentionLimits(EpubVersion version, bool metadataLimit) {
        EpubPublication original = EpubWritingContracts.CreateBook(version);
        byte[] rawPackage = Encoding.UTF8.GetBytes(original.GetPackageXml().ToString(SaveOptions.DisableFormatting));
        byte[] bytes = EpubWritingContracts.ReplaceEntry(original.Write().Bytes, original.PackagePath, rawPackage);
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
        long expanded = archive.Entries.Sum(entry => entry.Length);
        EpubPublication book = EpubPublication.Load(new MemoryStream(bytes), new EpubPublicationLoadOptions {
            MaxMetadataBytes = metadataLimit ? rawPackage.Length : 4L * 1024 * 1024,
            MaxExpandedBytes = metadataLimit ? 256L * 1024 * 1024 : expanded
        });
        EpubWriteResult result = book.Write();
        Assert.True(result.Report.UsedOriginalPackage);
        Assert.Equal(bytes, result.Bytes);
    }
}
