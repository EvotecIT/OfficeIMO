using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Epub.Benchmarks.Comparisons;

internal sealed record EpubComparisonScale(string Name, int Chapters, int ParagraphsPerChapter);

internal static class EpubComparisonCorpus {
    internal const string Title = "OfficeIMO EPUB benchmark";
    internal const string Creator = "OfficeIMO benchmark suite";
    internal const string Language = "en";

    internal static readonly IReadOnlyList<EpubComparisonScale> Scales = new[] {
        new EpubComparisonScale("Small", 8, 25),
        new EpubComparisonScale("Normal", 48, 80)
    };

    internal static EpubComparisonScale Get(string name) =>
        Scales.FirstOrDefault(scale => string.Equals(scale.Name, name, StringComparison.OrdinalIgnoreCase))
        ?? throw new ArgumentException($"Unknown scale '{name}'.", nameof(name));

    internal static byte[] CreatePackage(EpubComparisonScale scale) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteEntry(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
            WriteEntry(archive, "META-INF/container.xml", ContainerXml);
            WriteEntry(archive, "EPUB/package.opf", CreatePackageDocument(scale));
            WriteEntry(archive, "EPUB/nav.xhtml", CreateNavigationDocument(scale));
            for (var chapter = 0; chapter < scale.Chapters; chapter++) {
                WriteEntry(archive, ChapterPath(chapter), CreateChapterDocument(scale, chapter));
            }
        }

        return output.ToArray();
    }

    internal static string ChapterPath(int chapter) => $"EPUB/chapter-{chapter + 1:D4}.xhtml";

    internal static string ChapterMarker(int chapter, int paragraph) =>
        $"chapter-{chapter + 1:D4}-paragraph-{paragraph + 1:D4}";

    private static string CreatePackageDocument(EpubComparisonScale scale) {
        var manifest = new StringBuilder();
        var spine = new StringBuilder();
        manifest.AppendLine("    <item id=\"nav\" href=\"nav.xhtml\" media-type=\"application/xhtml+xml\" properties=\"nav\" />");
        for (var chapter = 0; chapter < scale.Chapters; chapter++) {
            manifest.Append("    <item id=\"chapter-")
                .Append(chapter + 1)
                .Append("\" href=\"chapter-")
                .Append((chapter + 1).ToString("D4"))
                .AppendLine(".xhtml\" media-type=\"application/xhtml+xml\" />");
            spine.Append("    <itemref idref=\"chapter-")
                .Append(chapter + 1)
                .AppendLine("\" />");
        }

        return $"""
            <?xml version="1.0" encoding="utf-8"?>
            <package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="book-id">
              <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
                <dc:identifier id="book-id">urn:uuid:officeimo-epub-benchmark</dc:identifier>
                <dc:title>{Title}</dc:title>
                <dc:creator>{Creator}</dc:creator>
                <dc:language>{Language}</dc:language>
                <meta property="dcterms:modified">2020-01-01T00:00:00Z</meta>
              </metadata>
              <manifest>
            {manifest}  </manifest>
              <spine>
            {spine}  </spine>
            </package>
            """;
    }

    private static string CreateNavigationDocument(EpubComparisonScale scale) {
        var items = new StringBuilder();
        for (var chapter = 0; chapter < scale.Chapters; chapter++) {
            items.Append("        <li><a href=\"chapter-")
                .Append((chapter + 1).ToString("D4"))
                .Append(".xhtml\">Chapter ")
                .Append(chapter + 1)
                .AppendLine("</a></li>");
        }

        return $"""
            <?xml version="1.0" encoding="utf-8"?>
            <html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" lang="en">
              <head><title>Contents</title></head>
              <body>
                <nav epub:type="toc" id="toc">
                  <h1>Contents</h1>
                  <ol>
            {items}      </ol>
                </nav>
              </body>
            </html>
            """;
    }

    private static string CreateChapterDocument(EpubComparisonScale scale, int chapter) {
        var paragraphs = new StringBuilder();
        for (var paragraph = 0; paragraph < scale.ParagraphsPerChapter; paragraph++) {
            string marker = ChapterMarker(chapter, paragraph);
            paragraphs.Append("    <p data-sequence=\"")
                .Append(marker)
                .Append("\">")
                .Append(marker)
                .Append(" carries deterministic benchmark prose for section ")
                .Append((chapter * scale.ParagraphsPerChapter) + paragraph + 1)
                .AppendLine(" with entities &amp; UTF-8 text: Zażółć gęślą jaźń.</p>");
        }

        return $"""
            <?xml version="1.0" encoding="utf-8"?>
            <html xmlns="http://www.w3.org/1999/xhtml" lang="en">
              <head><title>Chapter {chapter + 1}</title></head>
              <body>
                <h1>Chapter {chapter + 1}</h1>
            {paragraphs}  </body>
            </html>
            """;
    }

    private static void WriteEntry(
        ZipArchive archive,
        string path,
        string content,
        CompressionLevel compressionLevel = CompressionLevel.Optimal) {
        ZipArchiveEntry entry = archive.CreateEntry(path, compressionLevel);
        entry.LastWriteTime = new DateTimeOffset(2020, 1, 1, 0, 0, 0, TimeSpan.Zero);
        using Stream stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(content);
    }

    private const string ContainerXml = """
        <?xml version="1.0" encoding="utf-8"?>
        <container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
          <rootfiles>
            <rootfile full-path="EPUB/package.opf" media-type="application/oebps-package+xml" />
          </rootfiles>
        </container>
        """;
}
