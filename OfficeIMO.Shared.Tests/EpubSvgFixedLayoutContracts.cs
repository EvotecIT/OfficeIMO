using System.Threading;
using OfficeIMO.Epub;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubSvgFixedLayoutContracts {
    [Fact]
    public void SvgCanvasRoundTripsWithoutRewritingArtworkOrSemanticLinks() {
        var book = Book();
        var original = book.GetContentXml("page");
        var position = book.Spine.Single();
        position.Properties = "page-spread-left rendition:align-x-center";
        book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600) {
            Orientation = EpubPageOrientation.Landscape, Spread = EpubPageSpread.None, Side = EpubPageSide.Center
        });
        Assert.Equal("rendition:align-x-center rendition:layout-pre-paginated rendition:orientation-landscape rendition:spread-none rendition:page-spread-center", position.Properties);
        var reopened = EpubPublication.Load(new MemoryStream(book.Write().Bytes));
        var root = reopened.GetContentXml("page").Root!;
        Assert.Equal("0 0 800 600", (string?)root.Attribute("viewBox"));
        Assert.Equal("800", (string?)root.Attribute("width"));
        Assert.Equal("600", (string?)root.Attribute("height"));
        Assert.Equal("xMinYMin meet", (string?)root.Attribute("preserveAspectRatio"));
        Assert.Equal(original.Root!.Elements().Select(e => e.ToString()), root.Elements().Select(e => e.ToString()));
        Assert.Equal("title description", (string?)root.Attribute("aria-labelledby"));
        Assert.Single(reopened.Read().TableOfContents);
        reopened.SetFixedLayoutPage("page", new EpubFixedLayoutPage(600, 800));
        Assert.Equal("0 0 600 800", (string?)reopened.GetContentXml("page").Root!.Attribute("viewBox"));
        Assert.DoesNotContain("page-spread-", reopened.Spine.Single().Properties!);
    }

    [Fact]
    public void SvgRejectsXhtmlRegionsAndInvalidContentWithoutPartialPackageEdits() {
        var book = Book(); byte[] before = book.Write().Bytes;
        Assert.Throws<NotSupportedException>(() => book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600) {
            Regions = new[] { new EpubFixedLayoutRegion("art", 0, 0, 100, 100) }
        }));
        Assert.Equal(before, book.Write().Bytes);
        var xml = book.GetContentXml("page");
        xml.Root!.Add(new XElement(xml.Root.Name.Namespace + "rect", new XAttribute("id", "art")));
        book.SetContentXml("page", xml);
        byte[] invalidContent = book.GetResourceBytes("page");
        string package = book.GetPackageXml().ToString();
        Assert.Throws<InvalidDataException>(() => book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600)));
        Assert.Equal(invalidContent, book.GetResourceBytes("page"));
        Assert.Equal(package, book.GetPackageXml().ToString());
    }

    [Fact]
    public void SvgCanvasBudgetAndCancellationAreAtomic() {
        var book = Book(); byte[] input = book.Write().Bytes;
        using var zip = new ZipArchive(new MemoryStream(input), ZipArchiveMode.Read);
        var limited = EpubPublication.Load(new MemoryStream(input), new EpubPublicationLoadOptions { MaxExpandedBytes = zip.Entries.Sum(e => e.Length) });
        Assert.Throws<InvalidDataException>(() => limited.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600)));
        Assert.Equal(input, limited.Write().Bytes);
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        Assert.Throws<OperationCanceledException>(() => book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(800, 600), cancel.Token));
        Assert.Equal(input, book.Write().Bytes);
    }

    private static EpubPublication Book() {
        var book = EpubPublication.Create("SVG page", "en", "urn:example:svg-page");
        book.AddResource("page", "EPUB/page.svg", "image/svg+xml", Encoding.UTF8.GetBytes("""
            <svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" version="1.1"
                 width="100%" height="100%" viewBox="-10 -20 400 300" preserveAspectRatio="xMinYMin meet" aria-labelledby="title description">
              <title id="title">SVG page</title><desc id="description">A blue rectangle with a linked label.</desc>
              <rect id="art" x="20" y="20" width="200" height="100" fill="blue"/>
              <a xlink:href="#art"><text x="20" y="150">Read the diagram</text></a>
            </svg>
            """));
        book.AddSpineItem("page");
        book.SetNavigation(new[] { new EpubNavigationEntry("SVG page", "EPUB/page.svg") });
        return book;
    }
}
