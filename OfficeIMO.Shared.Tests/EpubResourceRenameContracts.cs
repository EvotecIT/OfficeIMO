using OfficeIMO.Epub;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubResourceRenameContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=");

    [Theory]
    [InlineData(EpubVersion.Epub2)]
    [InlineData(EpubVersion.Epub3)]
    public void ChapterAndNavigationMovesRepairBothDirectionsAndRetainManifestIdentity(EpubVersion version) {
        var book = Book(version);
        var declaration = book.Manifest.Single(item => item.Id == "one");
        book.RenameResource("one", "EPUB/parts/first chapter.xhtml");
        Assert.Equal("EPUB/parts/first chapter.xhtml", declaration.Reference.ContainerPath);
        Assert.Equal("../text/two.xhtml?view=1#two", (string?)book.GetContentXml("one").Descendants(Html + "a").Single().Attribute("href"));
        Assert.Equal("../parts/first%20chapter.xhtml#one", (string?)book.GetContentXml("two").Descendants(Html + "a").Single().Attribute("href"));
        book.RenameResource("navigation", version == EpubVersion.Epub3 ? "EPUB/navigation/toc.xhtml" : "EPUB/navigation/toc.ncx");
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal(new[] { "one", "two" }, reopened.Spine.Select(item => item.ManifestId));
        Assert.Equal("EPUB/parts/first chapter.xhtml", reopened.Read().TableOfContents[0].Target);
        Assert.DoesNotContain("EPUB/text/one.xhtml", reopened.EntryPaths);
        Assert.Equal(Png, reopened.GetResourceBytes("image"));
    }

    [Fact]
    public void ImageMoveRepairsInactiveCssImportsSrcsetAndSvgWithoutTouchingUnrelatedBytes() {
        var book = Book();
        book.AddResource("style", "EPUB/styles/main.css", "text/css", Encoding.UTF8.GetBytes("/* retain if untouched */ @import 'child.css' print; @media print { p{background:url('../images/dot.png')} }"));
        book.AddResource("child", "EPUB/styles/child.css", "text/css", Encoding.UTF8.GetBytes("p{background-image:image-set('../images/dot.png' 1x)}"));
        var xml = book.GetContentXml("one");
        xml.Root!.Element(Html + "head")!.Add(new XElement(Html + "link", new XAttribute("rel", "stylesheet"), new XAttribute("href", "../styles/main.css")));
        xml.Root.Element(Html + "body")!.Add(XElement.Parse("<svg xmlns='http://www.w3.org/2000/svg'><image href='../images/dot.png'/></svg>"),
            new XElement(Html + "img", new XAttribute("alt", "Dot"), new XAttribute("src", "../images/dot.png"), new XAttribute("srcset", "../images/dot.png  1x, ../images/dot.png 2x")));
        book.SetContentXml("one", xml);
        byte[] untouched = book.GetResourceBytes("two");
        book.RenameResource("image", "EPUB/art/dot.png");
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        Assert.Equal(untouched, reopened.GetResourceBytes("two"));
        Assert.Contains("../art/dot.png", Encoding.UTF8.GetString(reopened.GetResourceBytes("style")));
        Assert.Contains("../art/dot.png", Encoding.UTF8.GetString(reopened.GetResourceBytes("child")));
        Assert.Equal("../art/dot.png  1x, ../art/dot.png 2x", (string?)reopened.GetContentXml("one").Descendants(Html + "img").Single().Attribute("srcset"));
        Assert.Equal("../art/dot.png", (string?)reopened.GetContentXml("one").Descendants(XName.Get("image", "http://www.w3.org/2000/svg")).Single().Attribute("href"));
        book.RenameResource("style", "EPUB/css/deep/main.css");
        Assert.Contains("../../styles/child.css", Encoding.UTF8.GetString(book.GetResourceBytes("style")));
        book.Write();
    }

    [Fact]
    public void MovedHtmlBaseRetainsItsOriginalResolution() {
        var book = Book(); var xml = book.GetContentXml("one");
        xml.Root!.Element(Html + "head")!.Add(new XElement(Html + "base", new XAttribute("href", "../text/")));
        book.SetContentXml("one", xml);
        book.RenameResource("one", "EPUB/parts/deep/one.xhtml");
        var content = book.GetContentXml("one");
        string baseHref = content.Descendants(Html + "base").Single().Attribute("href")!.Value;
        Assert.Equal("../../text/", baseHref);
        Assert.Equal("EPUB/text/two.xhtml", EpubReference.Resolve("EPUB/parts/deep/one.xhtml", baseHref, content.Descendants(Html + "a").Single().Attribute("href")!.Value).ContainerPath);
        book.Write();
    }

    [Theory]
    [InlineData("collision")]
    [InlineData("opaque")]
    [InlineData("xml-base")]
    [InlineData("animation")]
    [InlineData("refresh")]
    public void UnsupportedOrCollidingRenameLeavesAllBytesAndDeclarationsUnchanged(string failure) {
        var book = Book();
        if (failure == "opaque") book.AddResource("opaque", "EPUB/opaque.bin", "application/octet-stream", new byte[] { 1 });
        if (failure == "xml-base") {
            var xml = book.GetContentXml("one"); xml.Root!.SetAttributeValue(XNamespace.Xml + "base", "text/"); book.SetContentXml("one", xml);
        }
        if (failure == "animation" || failure == "refresh") {
            var xml = book.GetContentXml("one");
            if (failure == "animation") xml.Root!.Element(Html + "body")!.Add(XElement.Parse("<svg xmlns='http://www.w3.org/2000/svg'><image id='drawing' href='../images/dot.png'/><animate href='#drawing' attributeName='href' values='../images/dot.png;../images/dot.png' dur='1s'/></svg>"));
            else xml.Root!.Element(Html + "head")!.Add(new XElement(Html + "meta", new XAttribute("http-equiv", "refresh"), new XAttribute("content", "2;url=two.xhtml")));
            book.SetContentXml("one", xml);
        }
        byte[] before = book.Write().Bytes;
        Assert.ThrowsAny<Exception>(() => book.RenameResource("one", failure == "collision" ? "EPUB/text/two.xhtml" : "EPUB/renamed.xhtml"));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void PreCancelledRenameDoesNotChangePublication() {
        var book = Book(); byte[] before = book.Write().Bytes;
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => book.RenameResource("one", "EPUB/new.xhtml", cancellation.Token));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void RenameChainsReportRelocationWhileLaterDeletionStillReportsLoss() {
        var book = EpubPublication.Load(new MemoryStream(Book().Write().Bytes));
        book.RenameResource("image", "EPUB/art/first.png");
        book.RenameResource("image", "EPUB/art/final.png");
        var report = book.Write().Report;
        Assert.Equal("EPUB/art/final.png", report.RenamedEntries["EPUB/images/dot.png"]);
        report.RequireNoLoss();
        book.RemoveResource("image");
        var removed = book.Write().Report;
        Assert.Empty(removed.RenamedEntries);
        Assert.Contains(removed.FidelityDiagnostics, item => item.Code == "EPUB_WRITE_ENTRY_REMOVED" && item.Location == "EPUB/images/dot.png");
    }

    [Fact]
    public void SmilTextAudioAndTextrefReferencesFollowRenamedResources() {
        var book = Book();
        book.AddResource("audio", "EPUB/audio/read.mp3", "audio/mpeg", new byte[] { 1, 2 });
        book.AddResource("overlay", "EPUB/overlays/one.smil", "application/smil+xml", Encoding.UTF8.GetBytes(
            "<smil xmlns='http://www.w3.org/ns/SMIL' xmlns:epub='http://www.idpf.org/2007/ops' version='3.0'><body><seq epub:textref='../text/one.xhtml'><par><text src='../text/one.xhtml#one'/><audio src='../audio/read.mp3' clipBegin='0s' clipEnd='1s'/></par></seq></body></smil>"));
        book.Manifest.Single(item => item.Id == "one").MediaOverlayId = "overlay";
        book.RenameResource("one", "EPUB/renamed/one.xhtml");
        book.RenameResource("audio", "EPUB/renamed/read.mp3");
        book.RenameResource("overlay", "EPUB/renamed/one.smil");
        XNamespace smil = "http://www.w3.org/ns/SMIL";
        var xml = book.GetContentXml("overlay");
        EpubReference Resolve(XElement element, XName attribute) => EpubReference.Resolve("EPUB/renamed/one.smil", element.Attribute(attribute)!.Value);
        var text = Resolve(xml.Descendants(smil + "text").Single(), "src");
        Assert.Equal("EPUB/renamed/one.xhtml", text.ContainerPath); Assert.Equal("one", text.Fragment);
        Assert.Equal("EPUB/renamed/one.xhtml", Resolve(xml.Descendants(smil + "seq").Single(), XName.Get("textref", "http://www.idpf.org/2007/ops")).ContainerPath);
        Assert.Equal("EPUB/renamed/read.mp3", Resolve(xml.Descendants(smil + "audio").Single(), "src").ContainerPath);
        Assert.Equal("1s", (string?)xml.Descendants(smil + "audio").Single().Attribute("clipEnd"));
        book.Write();
    }

    [Fact]
    public void RetentionFailureLeavesTheLoadedPublicationByteIdentical() {
        byte[] bytes = Book().Write().Bytes;
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
        long maximum = archive.Entries.Max(entry => entry.Length);
        var book = EpubPublication.Load(new MemoryStream(bytes), new EpubPublicationLoadOptions { MaxEntryBytes = maximum });
        Assert.Throws<InvalidDataException>(() => book.RenameResource("one", "EPUB/" + new string('a', 240) + "/" + new string('b', 240) + "/one.xhtml"));
        Assert.Equal(bytes, book.Write().Bytes);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReusingOriginalPathCannotReassignOriginalResourceIdentity(bool renameReplacement) {
        var book = EpubPublication.Load(new MemoryStream(Book().Write().Bytes));
        book.RenameResource("image", "EPUB/art/original.png");
        book.AddResource("new-image", "EPUB/images/dot.png", "image/png", Png);
        if (renameReplacement) book.RenameResource("new-image", "EPUB/art/new.png");
        Assert.Equal("EPUB/art/original.png", book.Write().Report.RenamedEntries["EPUB/images/dot.png"]);
        book.RemoveResource("image");
        var report = book.Write().Report;
        Assert.Empty(report.RenamedEntries);
        Assert.True(report.HasLoss);
        Assert.Contains(report.FidelityDiagnostics, item => item.Code == "EPUB_WRITE_ENTRY_REMOVED" && item.Location == "EPUB/images/dot.png");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void XmlStylesheetProcessingInstructionFollowsStylesheetAndOwnerMoves(bool moveOwner) {
        var book = Book();
        book.AddResource("style", "EPUB/styles/diagram.css", "text/css", Encoding.UTF8.GetBytes("rect{fill:blue}"));
        book.AddResource("diagram", "EPUB/images/diagram.svg", "image/svg+xml", Encoding.UTF8.GetBytes(
            "<?xml-stylesheet type='text/css' href='../styles/diagram.css?x=1&amp;y=2' title='A  B&#x9;C' media='screen' alternate='no'?><svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1 1'><rect width='1' height='1'/></svg>"));
        if (moveOwner) book.RenameResource("diagram", "EPUB/deep/art/diagram.svg");
        else book.RenameResource("style", "EPUB/revised/diagram.css");
        string owner = book.Manifest.Single(item => item.Id == "diagram").Reference.ContainerPath!;
        string target = book.Manifest.Single(item => item.Id == "style").Reference.ContainerPath!;
        var instruction = book.GetContentXml("diagram").DescendantNodes().OfType<XProcessingInstruction>().Single();
        var attributes = XElement.Parse("<style " + instruction.Data + "/>");
        Assert.Equal(target, EpubReference.Resolve(owner, attributes.Attribute("href")!.Value).ContainerPath);
        Assert.Equal("x=1&y=2", EpubReference.Resolve(owner, attributes.Attribute("href")!.Value).Query);
        Assert.Contains("title='A  B&#x9;C'", instruction.Data);
        Assert.Equal("screen", (string?)attributes.Attribute("media"));
        Assert.Equal("no", (string?)attributes.Attribute("alternate"));
        book.Write();
    }

    [Fact]
    public void ReplacingDeletedOriginalAtTheSamePathDoesNotInventRelocationEvidence() {
        var book = EpubPublication.Load(new MemoryStream(Book().Write().Bytes));
        book.RemoveResource("image");
        book.AddResource("replacement", "EPUB/images/dot.png", "image/png", Png);
        book.RenameResource("replacement", "EPUB/art/replacement.png");
        var report = book.Write().Report;
        Assert.Empty(report.RenamedEntries);
        Assert.True(report.HasLoss);
    }

    [Theory]
    [InlineData("custom-resource", "href='other.css'")]
    [InlineData("xml-stylesheet", "href='first.css' href='second.css'")]
    public void UninspectableProcessingInstructionsRejectTheWholeRename(string name, string data) {
        var book = Book();
        var document = book.GetContentXml("one"); document.AddFirst(new XProcessingInstruction(name, data));
        book.SetContentXml("one", document);
        byte[] before = book.Write().Bytes;
        Assert.ThrowsAny<Exception>(() => book.RenameResource("one", "EPUB/new/one.xhtml"));
        Assert.Equal(before, book.Write().Bytes);
    }

    private static EpubPublication Book(EpubVersion version = EpubVersion.Epub3) {
        var book = EpubPublication.Create("Rename", "en", version: version);
        book.AddResource("image", "EPUB/images/dot.png", "image/png", Png);
        book.AddChapter("one", "EPUB/text/one.xhtml", "One", "<h1 id='one'>One</h1><p><a href='two.xhtml?view=1#two'>Two</a></p>");
        book.AddChapter("two", "EPUB/text/two.xhtml", "Two", "<h1 id='two'>Two</h1><p><a href='one.xhtml#one'>One</a></p>");
        return book;
    }
}
