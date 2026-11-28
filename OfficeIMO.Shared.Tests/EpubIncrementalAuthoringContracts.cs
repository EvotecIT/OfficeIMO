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

    [Fact]
    public void BatchCanReplaceAChapterWhoseRetainedBytesDoNotParse() {
        var book = EpubPublication.Create("Repair", "en", "urn:example:repair");
        book.AddChapter("one", "EPUB/one.xhtml", "One", "<p>One.</p>");
        book.AddChapter("two", "EPUB/two.xhtml", "Two", "<p>Two.</p>");
        // XHTML 1.1 chapters in the wild use DTD entities such as &nbsp; that a DTD-less parse rejects.
        var imported = EpubPublication.Load(new MemoryStream(Replace(book.Write(Stamp).Bytes, "EPUB/two.xhtml",
            "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.1//EN\" \"http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd\">" +
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Two</title></head><body><p>a&nbsp;b</p></body></html>")));
        var repaired = new XDocument(new XElement(Html + "html", new XElement(Html + "head", new XElement(Html + "title", "Two")),
            new XElement(Html + "body", new XElement(Html + "p", "Repaired."))));

        imported.SetContentXml(new Dictionary<string, XDocument> { ["two"] = repaired });

        Assert.Contains("Repaired.", Encoding.UTF8.GetString(imported.GetResourceBytes("two")));
    }

    [Fact]
    public void AppendedNavigationMatchesWhenTwoPrefixesBindXhtml() {
        var book = EpubPublication.Create("Prefixes", "en", "urn:example:prefixes");
        book.AddChapter("c0", "EPUB/c0.xhtml", "Zero", "<p>Zero.</p>");
        var navigation = book.GetContentXml("navigation");
        navigation.Descendants(Html + "nav").Single().Add(new XAttribute(XNamespace.Xmlns + "h", Html.NamespaceName));
        book.SetContentXml("navigation", navigation);
        for (int index = 1; index < 4; index++) book.AddChapter("c" + index, "EPUB/c" + index + ".xhtml", "Chapter " + index, "<p>Text.</p>");
        var reloaded = EpubPublication.Load(new MemoryStream(book.Write(Stamp).Bytes));

        book.AddChapter("c4", "EPUB/c4.xhtml", "Four", "<p>Four.</p>");
        reloaded.AddChapter("c4", "EPUB/c4.xhtml", "Four", "<p>Four.</p>");

        Assert.Equal(reloaded.GetResourceBytes("navigation"), book.GetResourceBytes("navigation"));
    }

    [Fact]
    public void ContentPropertiesWriteWillAddCountTowardTheCallBudget() {
        var book = EpubPublication.Create("Math", "en", "urn:example:math");
        book.AddChapter("c0", "EPUB/c0.xhtml", "Zero", "<p>Zero.</p>");
        byte[] input = book.Write(Stamp).Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var limited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxMetadataBytes = zip.GetEntry("EPUB/package.opf")!.Length + 10 });
        XDocument content = MathChapter("Zero");

        // Write would declare properties="mathml", which doesn't fit in ten bytes.
        Assert.Throws<InvalidDataException>(() => limited.SetContentXml(new Dictionary<string, XDocument> { ["c0"] = content }));
        Assert.Equal(input, limited.Write(Stamp).Bytes);
    }

    [Fact]
    public void PropertiesFromEarlierEditsStayReservedForLaterOnes() {
        var book = EpubPublication.Create("Math", "en", "urn:example:math");
        book.AddChapter("c0", "EPUB/c0.xhtml", "Zero", "<p>Zero.</p>");
        book.AddChapter("c1", "EPUB/c1.xhtml", "One", "<p>One.</p>");
        byte[] input = book.Write(Stamp).Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        // Room for one properties="mathml" declaration (20 bytes), not two.
        var limited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxMetadataBytes = zip.GetEntry("EPUB/package.opf")!.Length + 25 });

        limited.SetContentXml(new Dictionary<string, XDocument> { ["c0"] = MathChapter("Zero") });
        Assert.Throws<InvalidDataException>(() => limited.SetContentXml(new Dictionary<string, XDocument> { ["c1"] = MathChapter("One") }));
        limited.Write(Stamp);
    }

    [Fact]
    public void MovingPropertiesBetweenDocumentsInOneBatchNeedsNoExtraRoom() {
        var book = EpubPublication.Create("Math", "en", "urn:example:math");
        book.AddChapter("c0", "EPUB/c0.xhtml", "Zero", "<p>Zero.</p>");
        book.AddChapter("c1", "EPUB/c1.xhtml", "One", "<p>One.</p>");
        book.SetContentXml(new Dictionary<string, XDocument> { ["c0"] = MathChapter("Zero") });
        byte[] input = book.Write(Stamp).Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var limited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxMetadataBytes = zip.GetEntry("EPUB/package.opf")!.Length });

        // Write drops properties="mathml" from c0 and adds it to c1, so the package size doesn't change.
        limited.SetContentXml(new Dictionary<string, XDocument> { ["c0"] = PlainChapter("Zero"), ["c1"] = MathChapter("One") });
        limited.Write(Stamp);
    }

    [Fact]
    public void RestoringAnImportedDocumentReleasesItsReservation() {
        var book = EpubPublication.Create("Math", "en", "urn:example:math");
        book.AddChapter("c0", "EPUB/c0.xhtml", "Zero", "<p>Zero.</p>");
        book.AddChapter("c1", "EPUB/c1.xhtml", "One", "<p>One.</p>");
        byte[] input = book.Write(Stamp).Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var limited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxMetadataBytes = zip.GetEntry("EPUB/package.opf")!.Length + 25 });
        XDocument original = limited.GetContentXml("c0");

        limited.SetContentXml(new Dictionary<string, XDocument> { ["c0"] = MathChapter("Zero") });
        limited.SetContentXml(new Dictionary<string, XDocument> { ["c0"] = original });
        limited.SetContentXml(new Dictionary<string, XDocument> { ["c1"] = MathChapter("One") });
        limited.Write(Stamp);
    }

    [Fact]
    public void ReplacingAPayloadOutsideABatchReleasesItsReservation() {
        var book = EpubPublication.Create("Math", "en", "urn:example:math");
        book.AddChapter("c0", "EPUB/c0.xhtml", "Zero", "<p>Zero.</p>");
        book.AddChapter("c1", "EPUB/c1.xhtml", "One", "<p>One.</p>");
        byte[] input = book.Write(Stamp).Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var limited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxMetadataBytes = zip.GetEntry("EPUB/package.opf")!.Length + 25 });

        limited.SetContentXml(new Dictionary<string, XDocument> { ["c0"] = MathChapter("Zero") });
        limited.SetContentXml("c0", PlainChapter("Zero"));
        limited.SetContentXml(new Dictionary<string, XDocument> { ["c1"] = MathChapter("One") });
        limited.Write(Stamp);
    }

    [Fact]
    public void PropertiesWriteWillRemoveDoNotMakeRoomForOtherEdits() {
        var book = EpubPublication.Create("Math", "en", "urn:example:math");
        book.AddChapter("c0", "EPUB/c0.xhtml", "Zero", "<p>Zero.</p>");
        book.SetContentXml(new Dictionary<string, XDocument> { ["c0"] = MathChapter("Zero") });
        byte[] input = book.Write(Stamp).Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var limited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxMetadataBytes = zip.GetEntry("EPUB/package.opf")!.Length });

        // Write measures the package before dropping properties="mathml", so the live package must fit on its own.
        limited.SetContentXml(new Dictionary<string, XDocument> { ["c0"] = PlainChapter("Zero") });
        Assert.Throws<InvalidDataException>(() => limited.Title = "Math!");
        limited.Write(Stamp);
    }

    [Fact]
    public void LineBreaksTheWriterEscapesCountTowardThePackageBudget() {
        var book = EpubPublication.Create("Breaks", "en", "urn:example:breaks");
        book.AddChapter("c0", "EPUB/c0.xhtml", "Zero", "<p>Zero.</p>");
        byte[] input = book.Write(Stamp).Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var limited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxMetadataBytes = zip.GetEntry("EPUB/package.opf")!.Length + 1500 });
        limited.AddDublinCoreMetadata("subject", "warm");

        // Each carriage return is written as &#xD;, so 500 of them need about 2,500 bytes.
        Assert.Throws<InvalidDataException>(() => limited.AddDublinCoreMetadata("description", "a" + new string('\r', 500) + "b"));
        limited.Write(Stamp);
    }

    [Fact]
    public void RejectedEditKeepsAnExplicitlyEmptyElement() {
        var book = EpubPublication.Create("Empty", "en", "urn:example:empty");
        book.AddChapter("c0", "EPUB/c0.xhtml", "Zero", "<p>Zero.</p>");
        byte[] written = book.Write(Stamp).Bytes;
        string package;
        using (var zip = new ZipArchive(new MemoryStream(written), ZipArchiveMode.Read))
        using (var reader = new StreamReader(zip.GetEntry("EPUB/package.opf")!.Open())) package = reader.ReadToEnd();
        package = System.Text.RegularExpressions.Regex.Replace(package, "<spine>.*</spine>", "<spine></spine>");
        var imported = EpubPublication.Load(new MemoryStream(Replace(written, "EPUB/package.opf", package)),
            new EpubPublicationLoadOptions { MaxMetadataBytes = Encoding.UTF8.GetByteCount(package) + 8 });
        string before = imported.GetPackageXml().ToString(SaveOptions.DisableFormatting);

        Assert.Throws<InvalidDataException>(() => imported.AddSpineItem("c0"));

        Assert.Equal(before, imported.GetPackageXml().ToString(SaveOptions.DisableFormatting));
        Assert.Contains("<spine></spine>", before);
    }

    private static XDocument MathChapter(string title) {
        XNamespace math = "http://www.w3.org/1998/Math/MathML";
        return new XDocument(new XElement(Html + "html", new XElement(Html + "head", new XElement(Html + "title", title)),
            new XElement(Html + "body", new XElement(math + "math", new XElement(math + "mi", "x")))));
    }

    private static XDocument PlainChapter(string title) =>
        new(new XElement(Html + "html", new XElement(Html + "head", new XElement(Html + "title", title)),
            new XElement(Html + "body", new XElement(Html + "p", title))));

    private static byte[] Replace(byte[] epub, string path, string content) {
        using var output = new MemoryStream();
        output.Write(epub, 0, epub.Length);
        using (var zip = new ZipArchive(output, ZipArchiveMode.Update, true)) {
            zip.GetEntry(path)!.Delete();
            using var writer = new StreamWriter(zip.CreateEntry(path).Open(), new UTF8Encoding(false));
            writer.Write(content);
        }
        return output.ToArray();
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
