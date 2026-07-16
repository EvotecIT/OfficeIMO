using OfficeIMO.Epub;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class EpubNavigationMetadataContractTests {
    [Fact]
    public void Load_ExposesEpub3RootfilesMetadataNavigationAndRemoteResources() {
        byte[] package = BuildEpub3Package();

        EpubDocument document = EpubDocument.Load(new MemoryStream(package, writable: false));

        Assert.Collection(
            document.Rootfiles,
            rootfile => {
                Assert.Equal("EPUB/missing.opf", rootfile.FullPath);
                Assert.False(rootfile.IsAvailable);
                Assert.False(rootfile.IsSelected);
            },
            rootfile => {
                Assert.Equal("EPUB/broken.opf", rootfile.FullPath);
                Assert.True(rootfile.IsAvailable);
                Assert.False(rootfile.IsSelected);
            },
            rootfile => {
                Assert.Equal("EPUB/package.opf", rootfile.FullPath);
                Assert.Equal("application/oebps-package+xml", rootfile.MediaType);
                Assert.True(rootfile.IsAvailable);
                Assert.True(rootfile.IsSelected);
            });
        Assert.Equal("EPUB/package.opf", document.OpfPath);
        Assert.Contains(document.Diagnostics, item => item.Code == "epub.container.multiple-rootfiles");
        Assert.Contains(document.Diagnostics, item => item.Code == "epub.package.rootfile-missing" && item.Path == "EPUB/missing.opf");
        Assert.Contains(document.Diagnostics, item => item.Code == "epub.package.invalid-xml" && item.Path == "EPUB/broken.opf");

        Assert.Equal(7, document.Metadata.Count);
        EpubMetadataEntry creator = Assert.Single(document.Metadata, item => item.Id == "creator");
        Assert.Equal(EpubMetadataKind.DublinCore, creator.Kind);
        Assert.Equal("Alice Example", creator.Value);
        Assert.Equal("aut", creator.Role);
        Assert.Equal("Example, Alice", creator.FileAs);
        EpubMetadataEntry role = Assert.Single(document.Metadata, item => item.Property == "role");
        Assert.Equal("#creator", role.Refines);
        Assert.Equal("marc:relators", role.Scheme);
        EpubMetadataEntry linkedRecord = Assert.Single(document.Metadata, item => item.Kind == EpubMetadataKind.Link);
        Assert.Equal("https://metadata.example/book.json", linkedRecord.Href);
        Assert.Equal("record", linkedRecord.Rel);

        Assert.Collection(
            document.TableOfContents,
            first => {
                Assert.Equal(EpubNavigationSource.Epub3Navigation, first.Source);
                Assert.Equal("Chapter One", first.Label);
                Assert.Equal("EPUB/chapters/one.xhtml", first.Target);
                Assert.Equal("intro", first.Fragment);
                EpubNavigationItem child = Assert.Single(first.Children);
                Assert.Equal("Chapter Two", child.Label);
                Assert.Equal("EPUB/chapters/two.xhtml", child.Target);
                Assert.Equal("details", child.Fragment);
            },
            missing => {
                Assert.Equal("Missing Appendix", missing.Label);
                Assert.Equal("EPUB/chapters/missing.xhtml", missing.Target);
            });
        EpubNavigationItem page = Assert.Single(document.PageList);
        Assert.Equal("1", page.Label);
        Assert.Equal("EPUB/chapters/one.xhtml", page.Target);
        Assert.Equal("page-1", page.Fragment);
        Assert.Collection(
            document.Landmarks,
            body => {
                Assert.Equal("Start reading", body.Label);
                Assert.Equal("bodymatter", body.SemanticType);
                Assert.False(body.IsRemote);
            },
            remote => {
                Assert.Equal("Publisher", remote.Label);
                Assert.Equal("https://publisher.example/book", remote.Target);
                Assert.True(remote.IsRemote);
            });

        EpubResource remoteCover = Assert.Single(document.Resources, item => item.Id == "remote-cover");
        Assert.True(remoteCover.IsRemote);
        Assert.Equal("https://cdn.example/cover.png", remoteCover.Href);
        Assert.Equal("https://cdn.example/cover.png", remoteCover.RemoteUri);
        Assert.Equal(0, remoteCover.LengthBytes);
        Assert.Null(remoteCover.Data);
        Assert.Contains(document.Diagnostics, item => item.Code == "epub.resource.remote" && item.Path == remoteCover.RemoteUri);
        Assert.Contains(document.Diagnostics, item => item.Code == "epub.navigation.remote-target");
        Assert.Contains(document.Diagnostics, item => item.Code == "epub.navigation.target-missing" && item.Path == "EPUB/chapters/missing.xhtml");
        Assert.Equal("NCX-only Chapter", document.Chapters[2].Title);
        Assert.DoesNotContain(document.TableOfContents, item => item.Target == "EPUB/chapters/three.xhtml");
    }

    [Fact]
    public void Load_ExposesEpub2NcxHierarchyPageListGuideAndLegacyMetadata() {
        byte[] package = BuildEpub2Package();

        EpubDocument document = EpubDocument.Load(new MemoryStream(package, writable: false));

        EpubNavigationItem part = Assert.Single(document.TableOfContents);
        Assert.Equal(EpubNavigationSource.Ncx, part.Source);
        Assert.Equal("Part One", part.Label);
        Assert.Equal(1, part.PlayOrder);
        EpubNavigationItem section = Assert.Single(part.Children);
        Assert.Equal("Section Two", section.Label);
        Assert.Equal(2, section.PlayOrder);
        Assert.Equal("OPS/two.xhtml", section.Target);
        Assert.Equal("section", section.Fragment);
        EpubNavigationItem page = Assert.Single(document.PageList);
        Assert.Equal("xii", page.Label);
        Assert.Equal("front", page.SemanticType);
        Assert.Equal(12, page.PlayOrder);
        EpubNavigationItem guide = Assert.Single(document.Landmarks);
        Assert.Equal(EpubNavigationSource.Epub2Guide, guide.Source);
        Assert.Equal("Cover", guide.Label);
        Assert.Equal("cover", guide.SemanticType);
        Assert.Equal("cover", guide.Fragment);

        Assert.Equal("Part One", document.Chapters[0].Title);
        Assert.Equal("Section Two", document.Chapters[1].Title);
        EpubMetadataEntry legacyCover = Assert.Single(document.Metadata, item => item.LegacyName == "cover");
        Assert.Equal("cover-image", legacyCover.Value);
        EpubMetadataEntry legacyCreator = Assert.Single(document.Metadata, item => item.Name == "creator");
        Assert.Equal("aut", legacyCreator.Role);
        Assert.Equal("Writer, Example", legacyCreator.FileAs);
    }

    [Fact]
    public void Load_BoundsMetadataNavigationCountAndDepthWithStableDiagnostics() {
        byte[] package = BuildEpub3Package();

        EpubDocument metadataLimited = EpubDocument.Load(
            new MemoryStream(package, writable: false),
            new EpubReadOptions { MaxMetadataItems = 2 });
        Assert.Equal(2, metadataLimited.Metadata.Count);
        Assert.Contains(metadataLimited.Diagnostics, item => item.Code == "epub.metadata.count-limit");

        EpubDocument countLimited = EpubDocument.Load(
            new MemoryStream(package, writable: false),
            new EpubReadOptions { MaxNavigationItems = 3 });
        Assert.Equal(3, CountNavigationItems(countLimited));
        Assert.Contains(countLimited.Diagnostics, item => item.Code == "epub.navigation.item-count-limit");

        EpubDocument depthLimited = EpubDocument.Load(
            new MemoryStream(package, writable: false),
            new EpubReadOptions { MaxNavigationDepth = 1 });
        Assert.Empty(depthLimited.TableOfContents[0].Children);
        Assert.Contains(depthLimited.Diagnostics, item => item.Code == "epub.navigation.depth-limit");
    }

    private static int CountNavigationItems(EpubDocument document) =>
        CountNavigationItems(document.TableOfContents) +
        CountNavigationItems(document.PageList) +
        CountNavigationItems(document.Landmarks);

    private static int CountNavigationItems(IEnumerable<EpubNavigationItem> items) {
        int count = 0;
        foreach (EpubNavigationItem item in items) {
            count++;
            count += CountNavigationItems(item.Children);
        }
        return count;
    }

    private static byte[] BuildEpub3Package() {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteTextEntry(
                archive,
                "META-INF/container.xml",
                "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\"><rootfiles>" +
                "<rootfile full-path=\"EPUB/missing.opf\" media-type=\"application/oebps-package+xml\"/>" +
                "<rootfile full-path=\"EPUB/broken.opf\" media-type=\"application/oebps-package+xml\"/>" +
                "<rootfile full-path=\"EPUB/package.opf\" media-type=\"application/oebps-package+xml\"/>" +
                "</rootfiles></container>");
            WriteTextEntry(archive, "EPUB/broken.opf", "<package><metadata>");
            WriteTextEntry(
                archive,
                "EPUB/package.opf",
                "<package version=\"3.0\" unique-identifier=\"book-id\" xmlns=\"http://www.idpf.org/2007/opf\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:opf=\"http://www.idpf.org/2007/opf\">" +
                "<metadata>" +
                "<dc:identifier id=\"book-id\">urn:navigation:book</dc:identifier>" +
                "<dc:title id=\"title\" xml:lang=\"en\">Navigation Book</dc:title>" +
                "<dc:creator id=\"creator\" opf:role=\"aut\" opf:file-as=\"Example, Alice\">Alice Example</dc:creator>" +
                "<meta refines=\"#creator\" property=\"role\" scheme=\"marc:relators\">aut</meta>" +
                "<meta property=\"dcterms:modified\">2026-07-16T00:00:00Z</meta>" +
                "<meta name=\"cover\" content=\"remote-cover\"/>" +
                "<link rel=\"record\" href=\"https://metadata.example/book.json\" media-type=\"application/json\"/>" +
                "</metadata><manifest>" +
                "<item id=\"nav\" href=\"nav.xhtml\" media-type=\"application/xhtml+xml\" properties=\"nav\"/>" +
                "<item id=\"ncx\" href=\"toc.ncx\" media-type=\"application/x-dtbncx+xml\"/>" +
                "<item id=\"one\" href=\"chapters/one.xhtml\" media-type=\"application/xhtml+xml\"/>" +
                "<item id=\"two\" href=\"chapters/two.xhtml\" media-type=\"application/xhtml+xml\"/>" +
                "<item id=\"three\" href=\"chapters/three.xhtml\" media-type=\"application/xhtml+xml\"/>" +
                "<item id=\"remote-cover\" href=\"https://cdn.example/cover.png\" media-type=\"image/png\" properties=\"cover-image\"/>" +
                "</manifest><spine toc=\"ncx\"><itemref idref=\"one\"/><itemref idref=\"two\"/><itemref idref=\"three\"/></spine></package>");
            WriteTextEntry(
                archive,
                "EPUB/nav.xhtml",
                "<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:epub=\"http://www.idpf.org/2007/ops\"><body>" +
                "<nav epub:type=\"toc\"><ol>" +
                "<li><a href=\"chapters/one.xhtml#intro\">Chapter One</a><ol>" +
                "<li><a href=\"chapters/two.xhtml#details\">Chapter Two</a></li></ol></li>" +
                "<li><a href=\"chapters/missing.xhtml\">Missing Appendix</a></li>" +
                "</ol></nav>" +
                "<nav epub:type=\"page-list\"><ol><li><a href=\"chapters/one.xhtml#page-1\">1</a></li></ol></nav>" +
                "<nav epub:type=\"landmarks\"><ol>" +
                "<li><a epub:type=\"bodymatter\" href=\"chapters/one.xhtml#intro\">Start reading</a></li>" +
                "<li><a epub:type=\"bibliography\" href=\"https://publisher.example/book\">Publisher</a></li>" +
                "</ol></nav></body></html>");
            WriteTextEntry(
                archive,
                "EPUB/toc.ncx",
                "<ncx xmlns=\"http://www.daisy.org/z3986/2005/ncx/\" version=\"2005-1\"><navMap>" +
                "<navPoint id=\"one\"><navLabel><text>NCX One</text></navLabel><content src=\"chapters/one.xhtml\"/></navPoint>" +
                "<navPoint id=\"three\"><navLabel><text>NCX-only Chapter</text></navLabel><content src=\"chapters/three.xhtml\"/></navPoint>" +
                "</navMap></ncx>");
            WriteTextEntry(archive, "EPUB/chapters/one.xhtml", "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><h1 id=\"intro\">One</h1><span id=\"page-1\">1</span></body></html>");
            WriteTextEntry(archive, "EPUB/chapters/two.xhtml", "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><h1 id=\"details\">Two</h1></body></html>");
            WriteTextEntry(archive, "EPUB/chapters/three.xhtml", "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><p>Third body</p></body></html>");
        }
        return output.ToArray();
    }

    private static byte[] BuildEpub2Package() {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteTextEntry(
                archive,
                "META-INF/container.xml",
                "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\"><rootfiles>" +
                "<rootfile full-path=\"OPS/content.opf\" media-type=\"application/oebps-package+xml\"/>" +
                "</rootfiles></container>");
            WriteTextEntry(
                archive,
                "OPS/content.opf",
                "<package version=\"2.0\" unique-identifier=\"book-id\" xmlns=\"http://www.idpf.org/2007/opf\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:opf=\"http://www.idpf.org/2007/opf\">" +
                "<metadata><dc:identifier id=\"book-id\">urn:epub2:navigation</dc:identifier><dc:title>EPUB 2 Navigation</dc:title>" +
                "<dc:creator opf:role=\"aut\" opf:file-as=\"Writer, Example\">Example Writer</dc:creator>" +
                "<meta name=\"cover\" content=\"cover-image\"/></metadata><manifest>" +
                "<item id=\"ncx\" href=\"toc.ncx\" media-type=\"application/x-dtbncx+xml\"/>" +
                "<item id=\"one\" href=\"one.xhtml\" media-type=\"application/xhtml+xml\"/>" +
                "<item id=\"two\" href=\"two.xhtml\" media-type=\"application/xhtml+xml\"/>" +
                "</manifest><spine toc=\"ncx\"><itemref idref=\"one\"/><itemref idref=\"two\"/></spine>" +
                "<guide><reference type=\"cover\" title=\"Cover\" href=\"one.xhtml#cover\"/></guide></package>");
            WriteTextEntry(
                archive,
                "OPS/toc.ncx",
                "<ncx xmlns=\"http://www.daisy.org/z3986/2005/ncx/\" version=\"2005-1\"><navMap>" +
                "<navPoint id=\"part\" playOrder=\"1\"><navLabel><text>Part One</text></navLabel><content src=\"one.xhtml\"/>" +
                "<navPoint id=\"section\" playOrder=\"2\"><navLabel><text>Section Two</text></navLabel><content src=\"two.xhtml#section\"/></navPoint>" +
                "</navPoint></navMap><pageList><pageTarget id=\"page-xii\" type=\"front\" value=\"xii\" playOrder=\"12\">" +
                "<navLabel><text>xii</text></navLabel><content src=\"one.xhtml#page-xii\"/></pageTarget></pageList></ncx>");
            WriteTextEntry(archive, "OPS/one.xhtml", "<html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Local One</title></head><body><h1 id=\"cover\">One</h1></body></html>");
            WriteTextEntry(archive, "OPS/two.xhtml", "<html xmlns=\"http://www.w3.org/1999/xhtml\"><head><title>Local Two</title></head><body><h1 id=\"section\">Two</h1></body></html>");
        }
        return output.ToArray();
    }

    private static void WriteTextEntry(ZipArchive archive, string path, string content) {
        ZipArchiveEntry entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using Stream stream = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes(content);
        stream.Write(bytes, 0, bytes.Length);
    }
}
