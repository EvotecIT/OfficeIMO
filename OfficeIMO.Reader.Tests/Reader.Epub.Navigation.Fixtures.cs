using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Tests;

public sealed partial class ReaderEpubModularTests {
    private static void BuildEpubWithNavigationMetadata(string epubPath) {
        using var file = new FileStream(epubPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: false);

        WriteNavigationEntry(
            archive,
            "META-INF/container.xml",
            "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\"><rootfiles>" +
            "<rootfile full-path=\"EPUB/missing.opf\" media-type=\"application/oebps-package+xml\"/>" +
            "<rootfile full-path=\"EPUB/package.opf\" media-type=\"application/oebps-package+xml\"/>" +
            "</rootfiles></container>");
        WriteNavigationEntry(
            archive,
            "EPUB/package.opf",
            "<package version=\"3.0\" unique-identifier=\"book-id\" xmlns=\"http://www.idpf.org/2007/opf\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:opf=\"http://www.idpf.org/2007/opf\">" +
            "<metadata>" +
            "<dc:identifier id=\"book-id\">urn:reader:navigation</dc:identifier>" +
            "<dc:title xml:lang=\"en\">Reader Navigation</dc:title>" +
            "<dc:creator id=\"creator\" opf:role=\"aut\">Reader Author</dc:creator>" +
            "<meta refines=\"#creator\" property=\"file-as\">Author, Reader</meta>" +
            "<meta property=\"dcterms:modified\">2026-07-16T00:00:00Z</meta>" +
            "</metadata><manifest>" +
            "<item id=\"nav\" href=\"nav.xhtml\" media-type=\"application/xhtml+xml\" properties=\"nav\"/>" +
            "<item id=\"one\" href=\"chapters/one.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "<item id=\"two\" href=\"chapters/two.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "<item id=\"local-cover\" href=\"images/cover.png\" media-type=\"image/png\"/>" +
            "<item id=\"remote-cover\" href=\"https://cdn.example/remote.png\" media-type=\"image/png\"/>" +
            "</manifest><spine><itemref idref=\"one\"/><itemref idref=\"two\"/></spine></package>");
        WriteNavigationEntry(
            archive,
            "EPUB/nav.xhtml",
            "<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:epub=\"http://www.idpf.org/2007/ops\"><body>" +
            "<nav epub:type=\"toc\"><ol><li><a href=\"chapters/one.xhtml#intro\">One</a><ol>" +
            "<li><a href=\"chapters/two.xhtml#details\">Two</a></li></ol></li></ol></nav>" +
            "<nav epub:type=\"page-list\"><ol><li><a href=\"chapters/one.xhtml#page-1\">1</a></li></ol></nav>" +
            "<nav epub:type=\"landmarks\"><ol>" +
            "<li><a epub:type=\"bodymatter\" href=\"chapters/one.xhtml#intro\">Start</a></li>" +
            "<li><a epub:type=\"bibliography\" href=\"https://publisher.example/book\">Publisher</a></li>" +
            "</ol></nav></body></html>");
        WriteNavigationEntry(
            archive,
            "EPUB/chapters/one.xhtml",
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><h1 id=\"intro\">One</h1><span id=\"page-1\">1</span>" +
            "<img src=\"../images/cover.png\" alt=\"Local cover\"/></body></html>");
        WriteNavigationEntry(
            archive,
            "EPUB/chapters/two.xhtml",
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><h1 id=\"details\">Two</h1>" +
            "<img src=\"https://cdn.example/remote.png\" alt=\"Remote cover\"/></body></html>");
        WriteNavigationBytes(archive, "EPUB/images/cover.png", new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 });
    }

    private static void WriteNavigationEntry(ZipArchive archive, string path, string content) =>
        WriteNavigationBytes(archive, path, Encoding.UTF8.GetBytes(content));

    private static void WriteNavigationBytes(ZipArchive archive, string path, byte[] content) {
        ZipArchiveEntry entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using Stream stream = entry.Open();
        stream.Write(content, 0, content.Length);
    }
}
