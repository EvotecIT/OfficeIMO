using OfficeIMO.Provenance;
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
        var archive = new List<KeyValuePair<string, byte[]>>();
        Add(archive, "mimetype", "application/epub+zip");
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
        return Archive(archive);
    }

    internal static byte[] OneChapter(string body, string extraMetadata = "", string packageAttributes = "", string spineProperties = "") =>
        Package(new[] { ("c", "chapter.xhtml", "application/xhtml+xml", "") },
            "<itemref idref='c' properties='" + spineProperties + "'/>",
            new[] { ("chapter.xhtml", Xhtml(body)) }, extraMetadata: extraMetadata, packageAttributes: packageAttributes);

    internal static byte[] ReplaceEntry(byte[] package, string path, byte[] bytes) {
        using var input = new MemoryStream(package);
        using var source = new ZipArchive(input, ZipArchiveMode.Read);
        var entries = source.Entries.Select(entry => {
            if (entry.FullName == path) return new KeyValuePair<string, byte[]>(entry.FullName, bytes);
            using var original = entry.Open();
            using var content = new MemoryStream();
            original.CopyTo(content);
            return new KeyValuePair<string, byte[]>(entry.FullName, content.ToArray());
        }).ToArray();
        return Archive(entries);
    }

    internal static byte[] Archive(IEnumerable<KeyValuePair<string, byte[]>> entries) {
        // Framework ZipArchive uses deflate even for NoCompression. EPUB requires
        // a physically stored leading mimetype entry on every target framework.
        var records = entries.OrderBy(entry => entry.Key == "mimetype" ? 0 : 1)
            .Select(entry => new OfficeProvenanceZipWriteEntry(entry.Key, entry.Value.LongLength,
                entry.Key != "mimetype", new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero), 0, 0,
                Array.Empty<byte>(), Array.Empty<byte>(), Array.Empty<byte>(),
                () => new MemoryStream(entry.Value, writable: false))).ToArray();
        return OfficeProvenanceZipWriter.Write(records, int.MaxValue);
    }

    private static void Add(List<KeyValuePair<string, byte[]>> entries, string path, string text) =>
        entries.Add(new KeyValuePair<string, byte[]>(path, new UTF8Encoding(false).GetBytes(text)));
}
