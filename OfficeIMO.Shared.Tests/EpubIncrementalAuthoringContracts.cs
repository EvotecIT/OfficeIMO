using OfficeIMO.Epub;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubIncrementalAuthoringContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";
    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=");
    private static readonly string LongId = "too-large-for-the-package-" + new string('x', 40);
    private static readonly EpubSaveOptions Stamp = new() { ModifiedAt = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero) };

    [Fact]
    public void RejectedFixedLayoutPageRestoresTheSpineAfterValidationFails() {
        var book = PageBook(12);
        var broken = book.GetContentXml("page_0007");
        broken.Root!.Element(Html + "body")!.Add(new XElement(Html + "img", new XAttribute("src", "../images/missing.png"), new XAttribute("alt", "")));
        book.SetContentXml("page_0007", broken);
        var position = book.Spine.Single(item => item.ManifestId == "page_0007");
        string? properties = position.Properties;
        string package = book.GetPackageXml().ToString();
        byte[] page = book.GetResourceBytes("page_0007");

        // The spine edit is applied before the edited page is validated, so this exercises the rollback.
        var error = Assert.Throws<InvalidDataException>(() => book.SetFixedLayoutPage("page_0007", new EpubFixedLayoutPage(800, 1200)));

        Assert.Contains("missing.png", error.Message);
        Assert.Equal(properties, position.Properties);
        Assert.Equal(package, book.GetPackageXml().ToString());
        Assert.Equal(page, book.GetResourceBytes("page_0007"));
        position.Properties = "rendition:layout-pre-paginated";
        Assert.Contains("rendition:layout-pre-paginated", book.GetPackageXml().ToString());
    }

    [Fact]
    public void RejectedAddChapterLeavesPackageNavigationAndIdsUnchanged() {
        byte[] input = PageBook(12).Write(Stamp).Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        long packageLength = zip.GetEntry("EPUB/package.opf")!.Length;
        var book = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxMetadataBytes = packageLength + 200 });
        book.AddChapter("small", "EPUB/text/small.xhtml", "Small", "<p>Fits.</p>");
        string package = book.GetPackageXml().ToString();
        byte[] navigation = book.GetResourceBytes("navigation");
        byte[] before = book.Write(Stamp).Bytes;

        Assert.Throws<InvalidDataException>(() => book.AddChapter(LongId, "EPUB/text/" + LongId + ".xhtml", "Too large", "<p>No room.</p>"));

        Assert.Equal(package, book.GetPackageXml().ToString());
        Assert.Equal(navigation, book.GetResourceBytes("navigation"));
        Assert.DoesNotContain("EPUB/text/" + LongId + ".xhtml", book.EntryPaths);
        Assert.Equal(before, book.Write(Stamp).Bytes);
        // Package ids are still tracked after the rollback.
        Assert.Throws<ArgumentException>(() => book.AddChapter("small", "EPUB/text/small-again.xhtml", "Again", "<p>Duplicate id.</p>"));
    }

    [Fact]
    public void ChapterNavigationMatchesAFreshlySerializedTableOfContents() {
        var book = PageBook(40);
        var reloaded = EpubPublication.Load(new MemoryStream(book.Write(Stamp).Bytes));

        // The first book appends to its cached table of contents; the reloaded one starts from its bytes.
        book.AddChapter("extra", "EPUB/text/extra.xhtml", "Extra & <more>", "<p>Extra.</p>");
        reloaded.AddChapter("extra", "EPUB/text/extra.xhtml", "Extra & <more>", "<p>Extra.</p>");

        Assert.Equal(reloaded.GetResourceBytes("navigation"), book.GetResourceBytes("navigation"));
        Assert.Equal(41, book.Read().TableOfContents.Count);
    }

    [Fact]
    public void LocalProblemsFailAtTheCallAndUnrelatedOnesAtSave() {
        var book = PageBook(4);
        var unrelated = book.GetContentXml("page_0001");
        unrelated.Root!.Element(Html + "body")!.Add(new XElement(Html + "img", new XAttribute("src", "../images/missing.png"), new XAttribute("alt", "")));
        book.SetContentXml("page_0001", unrelated);

        // Configuring another page validates that page, not the rest of the book.
        book.SetFixedLayoutPage("page_0003", new EpubFixedLayoutPage(600, 900));
        Assert.Throws<InvalidDataException>(() => book.Write(Stamp));

        // A link from the edited page is local, so a missing fragment fails at the call.
        var linking = book.GetContentXml("page_0002");
        linking.Root!.Element(Html + "body")!.Add(new XElement(Html + "a", new XAttribute("href", "page_0003.xhtml#nowhere"), "Next"));
        Assert.Throws<InvalidDataException>(() => book.SetContentXml(new Dictionary<string, XDocument> { ["page_0002"] = linking }));
    }

    [Fact]
    public void RemovingAnAnchorOtherPagesLinkToFailsAtTheCall() {
        var book = PageBook(3);
        var source = book.GetContentXml("page_0001");
        source.Root!.Element(Html + "body")!.Add(new XElement(Html + "a", new XAttribute("href", "page_0002.xhtml#art"), "Next"));
        book.SetContentXml(new Dictionary<string, XDocument> { ["page_0001"] = source });
        var target = book.GetContentXml("page_0002");
        target.Descendants().Single(element => (string?)element.Attribute("id") == "art").Attribute("id")!.Remove();
        byte[] before = book.Write(Stamp).Bytes;

        Assert.Throws<InvalidDataException>(() => book.SetContentXml(new Dictionary<string, XDocument> { ["page_0002"] = target }));
        Assert.Equal(before, book.Write(Stamp).Bytes);
    }

    // Builds a fixed-layout book one page at a time: chapter, image, then page canvas.
    internal static EpubPublication PageBook(int pages) {
        var book = EpubPublication.Create("Pages", "en", "urn:example:pages");
        book.AddStylesheet("style", "EPUB/style.css", "body { margin: 0; }");
        for (int index = 1; index <= pages; index++) AddPage(book, index);
        return book;
    }

    internal static void AddPage(EpubPublication book, int index) {
        string name = "page_" + index.ToString("D4");
        string image = "image-" + index.ToString("D4");
        book.AddChapter(name, "EPUB/text/" + name + ".xhtml", "Page " + index,
            "<div id='art'><img src='../images/" + image + ".png' alt='Page " + index + "' width='800' height='1200'/></div>", new[] { "style" });
        book.AddResource(image, "EPUB/images/" + image + ".png", "image/png", Png);
        book.SetFixedLayoutPage(name, new EpubFixedLayoutPage(800, 1200) { Spread = EpubPageSpread.None });
    }
}
