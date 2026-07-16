using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Tests;

public sealed partial class ReaderEpubModularTests {
    private static void BuildEpubWithResolvedResources(string epubPath) {
        using var file = new FileStream(epubPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create, leaveOpen: false);

        WriteResolvedResourceEntry(
            archive,
            "META-INF/container.xml",
            "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\"><rootfiles>" +
            "<rootfile full-path=\"EPUB/package.opf\" media-type=\"application/oebps-package+xml\"/>" +
            "</rootfiles></container>");
        WriteResolvedResourceEntry(
            archive,
            "EPUB/package.opf",
            "<package version=\"3.0\" xmlns=\"http://www.idpf.org/2007/opf\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\">" +
            "<metadata><dc:title>Resolved resources</dc:title></metadata><manifest>" +
            "<item id=\"chapter\" href=\"text/chapter.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "<item id=\"second\" href=\"text/second.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "<item id=\"cover\" href=\"shared/images/cover%20art.png\" media-type=\"image/png\"/>" +
            "<item id=\"reserved-image\" href=\"shared/images/cover%23v2.png\" media-type=\"image/png\"/>" +
            "<item id=\"root-image\" href=\"shared/images/root.png\" media-type=\"image/png\"/>" +
            "<item id=\"audio\" href=\"shared/audio/chapter.mp3\" media-type=\"audio/mpeg\"/>" +
            "<item id=\"video\" href=\"shared/video/clip.mp4\" media-type=\"video/mp4\"/>" +
            "<item id=\"styles\" href=\"shared/styles/book.css\" media-type=\"text/css\"/>" +
            "<item id=\"font\" href=\"shared/fonts/book.woff2\" media-type=\"font/woff2\"/>" +
            "<item id=\"remote-image\" href=\"https://cdn.example/remote.png\" media-type=\"image/png\"/>" +
            "</manifest><spine><itemref idref=\"chapter\"/><itemref idref=\"second\"/></spine></package>");
        WriteResolvedResourceEntry(
            archive,
            "EPUB/text/chapter.xhtml",
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><head><base href=\"../shared/\"/></head><body>" +
            "<h1 id=\"local\">Chapter</h1>" +
            "<a href=\"../text/chapter.xhtml?mode=print#local\">base link</a>" +
            "<img src=\"images/cover%20art.png?display=1#front\" alt=\"Cover\"/>" +
            "<img src=\"images/cover%23v2.png\" alt=\"Reserved name\"/>" +
            "<img src=\"/EPUB/shared/images/root.png\" alt=\"Root image\"/>" +
            "<img src=\"https://cdn.example/remote.png\" alt=\"Remote image\"/>" +
            "<img src=\"data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==\" alt=\"Inline\"/>" +
            "<audio><source src=\"audio/chapter.mp3\" type=\"audio/mpeg\"/>Audio fallback</audio>" +
            "<video src=\"video/clip.mp4#clip\">Video fallback</video>" +
            "</body></html>");
        WriteResolvedResourceEntry(
            archive,
            "EPUB/text/second.xhtml",
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body>" +
            "<h1 id=\"second\">Second</h1>" +
            "<a href=\"#second\">same fragment</a>" +
            "<a href=\"?view=print\">same query</a>" +
            "<a href=\"/EPUB/text/chapter.xhtml#local\">root link</a>" +
            "<a href=\"../../../outside.xhtml\">unsafe link</a>" +
            "<img src=\"../../../outside.png\" alt=\"Unsafe\"/>" +
            "</body></html>");

        WriteResolvedResourceBytes(archive, "EPUB/shared/images/cover art.png", new byte[] { 137, 80, 78, 71 });
        WriteResolvedResourceBytes(archive, "EPUB/shared/images/cover#v2.png", new byte[] { 137, 80, 78, 71, 2 });
        WriteResolvedResourceBytes(archive, "EPUB/shared/images/root.png", new byte[] { 137, 80, 78, 71, 1 });
        WriteResolvedResourceBytes(archive, "EPUB/shared/audio/chapter.mp3", new byte[] { 73, 68, 51, 4 });
        WriteResolvedResourceBytes(archive, "EPUB/shared/video/clip.mp4", new byte[] { 0, 0, 0, 20, 102, 116, 121, 112 });
        WriteResolvedResourceEntry(archive, "EPUB/shared/styles/book.css", "body { font-family: Book; }");
        WriteResolvedResourceBytes(archive, "EPUB/shared/fonts/book.woff2", new byte[] { 119, 79, 70, 50 });
    }

    private static void WriteResolvedResourceEntry(ZipArchive archive, string path, string content) =>
        WriteResolvedResourceBytes(archive, path, Encoding.UTF8.GetBytes(content));

    private static void WriteResolvedResourceBytes(ZipArchive archive, string path, byte[] content) {
        ZipArchiveEntry entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using Stream stream = entry.Open();
        stream.Write(content, 0, content.Length);
    }
}
