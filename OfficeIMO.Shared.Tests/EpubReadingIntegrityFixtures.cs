using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Shared.Tests;

internal static class EpubIntegrityFixtures {
    internal static string Xhtml(string body) =>
        "<html xmlns='http://www.w3.org/1999/xhtml'><head><title>Book</title></head><body>" + body + "</body></html>";

    internal static byte[] Package(
        (string Id, string Href, string Media, string Properties)[] manifest,
        string spine,
        (string Path, string Text)[] entries,
        string version = "3.0",
        string spineAttributes = "",
        string extraMetadata = "",
        string packageAttributes = "",
        string? nav = null) {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true)) {
            Add(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
            Add(archive, "META-INF/container.xml",
                "<container xmlns='urn:oasis:names:tc:opendocument:xmlns:container' version='1.0'><rootfiles>" +
                "<rootfile full-path='EPUB/package.opf' media-type='application/oebps-package+xml'/></rootfiles></container>");
            string items = string.Concat(manifest.Select(i =>
                $"<item id='{i.Id}' href='{i.Href}' media-type='{i.Media}' properties='{i.Properties}'/>"));
            if (version == "3.0") {
                items += "<item id='nav' href='nav.xhtml' media-type='application/xhtml+xml' properties='nav'/>";
                string links = string.Concat(manifest.Where(i => i.Media == "application/xhtml+xml" || i.Media == "image/svg+xml")
                    .Select(i => $"<li><a href='{i.Href}'>{i.Id}</a></li>"));
                Add(archive, "EPUB/nav.xhtml", nav ??
                    "<html xmlns='http://www.w3.org/1999/xhtml' xmlns:epub='http://www.idpf.org/2007/ops'>" +
                    "<head><title>Contents</title></head><body><nav epub:type='toc'><ol>" + links + "</ol></nav></body></html>");
            }
            Add(archive, "EPUB/package.opf",
                $"<package xmlns='http://www.idpf.org/2007/opf' xmlns:dc='http://purl.org/dc/elements/1.1/' version='{version}' unique-identifier='uid' {packageAttributes}>" +
                "<metadata><dc:identifier id='uid'>urn:book:integrity</dc:identifier><dc:title>Integrity book</dc:title><dc:language>en</dc:language>" +
                "<meta property='dcterms:modified'>2026-10-02T00:00:00Z</meta>" + extraMetadata +
                "</metadata><manifest>" + items + "</manifest><spine " + spineAttributes + ">" + spine + "</spine></package>");
            foreach (var entry in entries) Add(archive, "EPUB/" + entry.Path, entry.Text);
        }
        return stream.ToArray();
    }

    internal static byte[] OneChapter(string body, string extraMetadata = "", string packageAttributes = "", string spineProperties = "") =>
        Package(new[] { ("c", "chapter.xhtml", "application/xhtml+xml", "") },
            "<itemref idref='c' properties='" + spineProperties + "'/>",
            new[] { ("chapter.xhtml", Xhtml(body)) }, extraMetadata: extraMetadata, packageAttributes: packageAttributes);

    internal static byte[] ReplaceEntry(byte[] package, string path, byte[] bytes) {
        using var input = new MemoryStream(package);
        using var output = new MemoryStream();
        using (var source = new ZipArchive(input, ZipArchiveMode.Read))
        using (var target = new ZipArchive(output, ZipArchiveMode.Create, true)) {
            foreach (var entry in source.Entries) {
                using var destination = target.CreateEntry(entry.FullName,
                    entry.FullName == "mimetype" ? CompressionLevel.NoCompression : CompressionLevel.Optimal).Open();
                if (entry.FullName == path) destination.Write(bytes, 0, bytes.Length);
                else { using var original = entry.Open(); original.CopyTo(destination); }
            }
        }
        return output.ToArray();
    }

    private static void Add(ZipArchive archive, string path, string text, CompressionLevel level = CompressionLevel.Optimal) {
        using var writer = new StreamWriter(archive.CreateEntry(path, level).Open(), new UTF8Encoding(false));
        writer.Write(text);
    }
}
