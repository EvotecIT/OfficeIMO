using System.IO.Compression;

namespace OfficeIMO.Tests;

public sealed partial class ReaderEpubModularTests {
    private static void BuildEpubWithStructuredSemantics(string epubPath) {
        using var file = new FileStream(epubPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: false);

        WriteTextEntry(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
        WriteTextEntry(
            archive,
            "META-INF/container.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">" +
            "<rootfiles><rootfile full-path=\"OEBPS/package.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles>" +
            "</container>");
        WriteTextEntry(
            archive,
            "OEBPS/package.opf",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<package version=\"3.0\" unique-identifier=\"bookid\" xmlns=\"http://www.idpf.org/2007/opf\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\">" +
            "<metadata><dc:identifier id=\"bookid\">semantics-book</dc:identifier><dc:title>Structured semantics</dc:title><dc:language>en</dc:language></metadata>" +
            "<manifest><item id=\"chapter\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "<item id=\"cover\" href=\"images/cover.png\" media-type=\"image/png\"/></manifest>" +
            "<spine><itemref idref=\"chapter\"/></spine></package>");
        WriteTextEntry(
            archive,
            "OEBPS/chapter.xhtml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<html xmlns=\"http://www.w3.org/1999/xhtml\" xmlns:epub=\"http://www.idpf.org/2007/ops\"><head><title>Semantics</title></head><body>" +
            "<div id=\"detail\" role=\"heading\" aria-level=\"4\">Accessible section</div>" +
            "<blockquote><p>Quoted wisdom.</p></blockquote>" +
            "<ol type=\"A\" start=\"3\"><li>Third choice</li></ol>" +
            "<table><caption>Coverage</caption><tr><th>Measure</th><th>Value</th></tr><tr><td>Footnotes</td><td>Typed</td></tr></table>" +
            "<pre data-language=\"csharp\">Console.WriteLine(3);</pre>" +
            "<p id=\"ref-alpha\">Evidence<a epub:type=\"noteref\" role=\"doc-noteref\" href=\"#note-alpha\">1</a>.</p>" +
            "<aside epub:type=\"footnote\" role=\"doc-footnote\" id=\"note-alpha\"><p>Source <strong>detail</strong>.</p>" +
            "<a epub:type=\"backlink\" role=\"doc-backlink\" href=\"#ref-alpha\">return</a></aside>" +
            "<p><a href=\"#detail\" aria-label=\"Read details\"></a></p>" +
            "<img src=\"images/cover.png\" alt=\"\"/>" +
            "<figure><span id=\"cover-label\">Accessible cover</span><img src=\"images/cover.png\" aria-labelledby=\"cover-label\"/></figure>" +
            "</body></html>");
        WriteBinaryEntry(archive, "OEBPS/images/cover.png", new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 });
    }
}
