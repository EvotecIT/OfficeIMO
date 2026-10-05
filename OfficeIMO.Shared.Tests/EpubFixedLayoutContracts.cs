using OfficeIMO.Epub;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubFixedLayoutContracts {
    private static readonly XNamespace Html = "http://www.w3.org/1999/xhtml";

    [Fact]
    public void PageConfigurationPreservesSemanticsAndReplacesOnlyPresentationOverrides() {
        var book = Book();
        book.DeclareVocabularyPrefix("fx", "http://www.idpf.org/vocab/rendition/#");
        var position = book.Spine.Single();
        position.Properties = "fx:layout-reflowable fx:orientation-portrait fx:spread-portrait page-spread-left rendition:align-x-center";
        var body = new XElement(book.GetContentXml("page").Root!.Element(Html + "body")!);
        book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600) {
            Orientation = EpubPageOrientation.Landscape, Spread = EpubPageSpread.Both, Side = EpubPageSide.Right
        });
        Assert.Equal("rendition:align-x-center rendition:layout-pre-paginated rendition:orientation-landscape rendition:spread-both page-spread-right", position.Properties);
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        var content = reopened.GetContentXml("page");
        Assert.True(XNode.DeepEquals(body, content.Root!.Element(Html + "body")));
        Assert.Equal("width=800, height=600", (string?)content.Descendants(Html + "meta").Single(e => (string?)e.Attribute("name") == "viewport").Attribute("content"));
        Assert.Contains("width: 800px; height: 600px;", content.Descendants(Html + "style").Single().Value);
        Assert.Equal("page", reopened.Spine.Single().ManifestId);
        Assert.Single(reopened.Read().TableOfContents);
        reopened.SetFixedLayoutPage("page", new EpubFixedLayoutPage(600, 800) { Side = EpubPageSide.Center, Spread = EpubPageSpread.None });
        content = reopened.GetContentXml("page");
        Assert.Single(content.Descendants(Html + "style"));
        Assert.Contains("rendition:page-spread-center", reopened.Spine.Single().Properties!);
        reopened.SetFixedLayoutPage("page", new EpubFixedLayoutPage(600, 800));
        Assert.DoesNotContain("page-spread-", reopened.Spine.Single().Properties!);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void InvalidDimensionsDoNotMutateThePublication(int value) {
        var book = Book(); var before = book.Write().Bytes;
        Assert.Throws<ArgumentOutOfRangeException>(() => book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(value, 600)));
        Assert.Throws<ArgumentOutOfRangeException>(() => book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, value)));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void RejectedContentAndPackageBudgetsAreAtomic() {
        var book = Book(); byte[] input = book.Write().Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var limited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxExpandedBytes = zip.Entries.Sum(e => e.Length) + 100 });
        Assert.Throws<InvalidDataException>(() => limited.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600)));
        Assert.Equal(input, limited.Write().Bytes);
        var packageLimited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxMetadataBytes = zip.GetEntry("EPUB/package.opf")!.Length + 10 });
        Assert.Throws<InvalidDataException>(() => packageLimited.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600)));
        Assert.Equal(input, packageLimited.Write().Bytes);
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600), cancellation.Token));
        Assert.Equal(input, book.Write().Bytes);
    }

    [Theory]
    [InlineData("<meta name='viewport' content='width=100,height=100'/><meta name='viewport' content='width=200,height=200'/>")]
    [InlineData("<style id='officeimo-fixed-layout-canvas'>p { color: red; }</style>")]
    public void AmbiguousImportedDeclarationsAreNotSilentlyDeleted(string headFragment) {
        var book = Book(); var xml = book.GetContentXml("page");
        var wrapper = XElement.Parse("<head xmlns='" + Html + "'>" + headFragment + "</head>");
        xml.Root!.Element(Html + "head")!.Add(wrapper.Nodes()); book.SetContentXml("page", xml);
        byte[] before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600)));
        Assert.Equal(before, book.Write().Bytes);
    }

    [Fact]
    public void NonSpineAndEpub2InputsAreRejected() {
        var book = Book(); var before = book.Write().Bytes;
        Assert.Throws<InvalidDataException>(() => book.SetFixedLayoutPage("navigation", new EpubFixedLayoutPage(800, 600)));
        Assert.Equal(before, book.Write().Bytes);
        var epub2 = EpubPublication.Create("Legacy", version: EpubVersion.Epub2);
        epub2.AddChapter("page", "EPUB/page.xhtml", "Page", "<p>Text</p>"); before = epub2.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => epub2.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600)));
        Assert.Equal(before, epub2.Write().Bytes);
    }

    private static EpubPublication Book() {
        var book = EpubPublication.Create("Page canvas", "en", "urn:example:fixed-layout");
        book.AddChapter("page", "EPUB/page.xhtml", "A page", "<h1 id='title'>A page</h1><section aria-labelledby='title'><p id='first'>First in reading order.</p><p id='second'>Second, with a <a href='#first'>backlink</a>.</p></section>");
        return book;
    }
}
