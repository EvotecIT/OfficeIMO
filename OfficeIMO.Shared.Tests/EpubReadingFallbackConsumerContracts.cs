using System.Text;
using System.Xml.Linq;
using OfficeIMO.Epub;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubReadingFallbackConsumerContracts {
    [Theory]
    [InlineData("application/xhtml+xml", "3.0")]
    [InlineData("image/svg+xml", "3.0")]
    [InlineData("application/xhtml+xml", "2.0")]
    public void SharedFallback_RetainsEachPrimaryNavigationLabel(string mediaType, string version) {
        string content = mediaType == "image/svg+xml" ? "<svg xmlns='http://www.w3.org/2000/svg'><text>Fallback text</text></svg>" :
            "<html xmlns='http://www.w3.org/1999/xhtml'><head/><body><p>Fallback text</p></body></html>";
        string nav = "<html xmlns='http://www.w3.org/1999/xhtml' xmlns:epub='http://www.idpf.org/2007/ops'><head><title>Contents</title></head><body>" +
            "<nav epub:type='toc'><ol><li><a href='foreign.bin'>First primary title</a></li><li><a href='other.bin'>Second primary title</a></li></ol></nav></body></html>";
        string ncx = "<ncx xmlns='http://www.daisy.org/z3986/2005/ncx/' version='2005-1'><head/><docTitle><text>Book</text></docTitle><navMap>" +
            "<navPoint id='a' playOrder='1'><navLabel><text>First primary title</text></navLabel><content src='foreign.bin'/></navPoint>" +
            "<navPoint id='b' playOrder='2'><navLabel><text>Second primary title</text></navLabel><content src='other.bin'/></navPoint></navMap></ncx>";
        byte[] package = EpubIntegrityFixtures.Package(new[] {
            ("foreign", "foreign.bin", "application/vnd.example.foreign", ""), ("other", "other.bin", "application/vnd.example.foreign", ""),
            ("fallback", "fallback.xhtml", mediaType, ""), ("ncx", "toc.ncx", "application/x-dtbncx+xml", "")
        }, "<itemref idref='foreign'/><itemref idref='other'/>", new[] {
            ("foreign.bin", "opaque"), ("other.bin", "opaque"), ("fallback.xhtml", content), ("toc.ncx", ncx)
        }, version: version, spineAttributes: "toc='ncx'", nav: nav);
        package = EditPackage(package, root => {
            foreach (XElement item in root.Descendants(Opf + "item").Where(item => (string?)item.Attribute("id") is "foreign" or "other"))
                item.SetAttributeValue("fallback", "fallback");
        });
        EpubDocument read = EpubDocument.Load(new MemoryStream(package));
        Assert.Equal(new[] { "First primary title", "Second primary title" }, read.Chapters.Select(chapter => chapter.Title));
        Assert.All(read.Chapters, chapter => Assert.Equal("EPUB/fallback.xhtml", chapter.Path));
        Assert.True(read.ReadSummary.IsComplete);
    }

    private static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";
    private static byte[] EditPackage(byte[] package, Action<XElement> edit) {
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(package));
        using Stream stream = archive.GetEntry("EPUB/package.opf")!.Open();
        XDocument xml = XDocument.Load(stream); edit(xml.Root!);
        return EpubIntegrityFixtures.ReplaceEntry(package, "EPUB/package.opf", Encoding.UTF8.GetBytes(xml.ToString()));
    }
}
