using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.Xml;
using OfficeIMO.Reader.Zip;
using OfficeIMO.Epub;
using OfficeIMO.Zip;
using System.Globalization;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderGoldenFixtureTests {
    [Fact]
    public void ReaderGolden_Csv_Path() {
        var inputPath = GetInputPath("sample.csv");
        var chunks = DocumentReaderCsvExtensions.ReadCsv(
            inputPath,
            csvOptions: new CsvReadOptions {
                ChunkRows = 2,
                IncludeMarkdown = true
            }).ToList();

        AssertGolden("csv", BuildSnapshot(chunks));
    }

    [Fact]
    public void ReaderGolden_Json_Path() {
        var inputPath = GetInputPath("sample.json");
        var chunks = DocumentReaderJsonExtensions.ReadJson(
            inputPath,
            jsonOptions: new JsonReadOptions {
                ChunkRows = 4,
                MaxDepth = 16,
                IncludeMarkdown = true
            }).ToList();

        AssertGolden("json", BuildSnapshot(chunks));
    }

    [Fact]
    public void ReaderGolden_Xml_Path() {
        var inputPath = GetInputPath("sample.xml");
        var chunks = DocumentReaderXmlExtensions.ReadXml(
            inputPath,
            xmlOptions: new XmlReadOptions {
                ChunkRows = 4,
                IncludeMarkdown = true
            }).ToList();

        AssertGolden("xml", BuildSnapshot(chunks));
    }

    [Fact]
    public void ReaderGolden_Html_Path() {
        var inputPath = GetInputPath("sample.html");
        var chunks = DocumentReaderHtmlExtensions.ReadHtmlFile(
            inputPath,
            readerOptions: new ReaderOptions {
                MaxChars = 8_000
            }).ToList();

        AssertGolden("html", BuildSnapshot(chunks));
    }

    [Fact]
    public void ReaderGolden_Zip_Stream() {
        using var zipStream = BuildZipFixtureStream();
        var chunks = DocumentReaderZipExtensions.ReadZip(
            zipStream,
            sourceName: "golden.zip",
            readerOptions: new ReaderOptions {
                MaxChars = 8_000
            },
            zipOptions: new ZipTraversalOptions {
                DeterministicOrder = true
            },
            readerZipOptions: new ReaderZipOptions {
                ReadNestedZipEntries = false
            }).ToList();

        AssertGolden("zip", BuildSnapshot(chunks));
    }

    [Fact]
    public void ReaderGolden_Epub_Stream() {
        using var epubStream = BuildEpubFixtureStream();
        var chunks = DocumentReaderEpubExtensions.ReadEpub(
            epubStream,
            sourceName: "golden.epub",
            readerOptions: new ReaderOptions {
                MaxChars = 8_000
            },
            epubOptions: new EpubReadOptions {
                PreferSpineOrder = true
            }).ToList();

        AssertGolden("epub", BuildSnapshot(chunks));
    }

    private static string BuildSnapshot(IReadOnlyList<ReaderChunk> chunks) {
        var sb = new StringBuilder();
        sb.AppendLine("count=" + chunks.Count.ToString(CultureInfo.InvariantCulture));

        for (int i = 0; i < chunks.Count; i++) {
            var chunk = chunks[i];
            sb.AppendLine("[chunk:" + i.ToString(CultureInfo.InvariantCulture) + "]");
            sb.AppendLine("id=" + chunk.Id);
            sb.AppendLine("kind=" + chunk.Kind.ToString());
            sb.AppendLine("path=" + NormalizePath(chunk.Location.Path));
            sb.AppendLine("block=" + (chunk.Location.BlockIndex?.ToString(CultureInfo.InvariantCulture) ?? string.Empty));
            sb.AppendLine("sourceBlock=" + (chunk.Location.SourceBlockIndex?.ToString(CultureInfo.InvariantCulture) ?? string.Empty));
            sb.AppendLine("heading=" + (chunk.Location.HeadingPath ?? string.Empty));
            sb.AppendLine("text=" + NormalizeValue(chunk.Text));
            sb.AppendLine("markdown=" + NormalizeValue(chunk.Markdown));
            sb.AppendLine("warnings=" + NormalizeWarnings(chunk.Warnings));
            sb.AppendLine("tables=" + NormalizeTables(chunk.Tables));
        }

        return sb.ToString().TrimEnd();
    }

    private static string NormalizeWarnings(IReadOnlyList<string>? warnings) {
        if (warnings == null || warnings.Count == 0) return string.Empty;
        return string.Join(" || ", warnings.Select(NormalizeValue));
    }

    private static string NormalizeTables(IReadOnlyList<ReaderTable>? tables) {
        if (tables == null || tables.Count == 0) return string.Empty;

        var parts = new List<string>(tables.Count);
        for (int i = 0; i < tables.Count; i++) {
            var table = tables[i];
            var columns = string.Join(",", table.Columns);
            var rowCount = table.Rows?.Count ?? 0;
            parts.Add(string.Concat(
                "t",
                i.ToString(CultureInfo.InvariantCulture),
                ":",
                columns,
                "#",
                rowCount.ToString(CultureInfo.InvariantCulture),
                "#",
                table.TotalRowCount.ToString(CultureInfo.InvariantCulture),
                "#truncated=",
                table.Truncated ? "true" : "false"));
        }

        return string.Join(" ; ", parts);
    }

    private static string NormalizeValue(string? value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;

        var normalized = value!
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Trim();

        if (normalized.Length > 220) {
            normalized = normalized.Substring(0, 220);
        }

        return normalized.Replace("\n", "\\n");
    }

    private static string NormalizePath(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;

        var normalized = value!.Replace('\\', '/');
        var root = GetTestsProjectRoot().Replace('\\', '/');
        if (normalized.StartsWith(root, StringComparison.OrdinalIgnoreCase)) {
            normalized = normalized.Substring(root.Length).TrimStart('/');
        }

        return normalized;
    }

    private static void AssertGolden(string name, string actualSnapshot) {
        var expectedPath = GetExpectedPath(name);
        if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_GOLDEN"), "1", StringComparison.Ordinal)) {
            File.WriteAllText(expectedPath, actualSnapshot + Environment.NewLine, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            return;
        }

        if (!File.Exists(expectedPath)) {
            throw new FileNotFoundException(
                "Golden snapshot missing. Set OFFICEIMO_UPDATE_GOLDEN=1 and re-run this test to generate it.",
                expectedPath);
        }

        var expected = File.ReadAllText(expectedPath, Encoding.UTF8);
        Assert.Equal(NormalizeValue(expected), NormalizeValue(actualSnapshot));
    }

    private static string GetInputPath(string fileName) {
        return Path.Combine(GetTestsProjectRoot(), "ReaderGolden", "Inputs", fileName);
    }

    private static string GetExpectedPath(string fixtureName) {
        var root = GetTestsProjectRoot();
        var expectedFolder = Path.Combine(root, "ReaderGolden", "Expected");
        Directory.CreateDirectory(expectedFolder);

        return Path.Combine(expectedFolder, fixtureName + ".snapshot.txt");
    }

    private static string GetTestsProjectRoot() {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null) {
            var candidate = Path.Combine(dir.FullName, "OfficeIMO.Tests.csproj");
            if (File.Exists(candidate)) {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate OfficeIMO.Tests project root from test runtime base directory.");
    }

    private static MemoryStream BuildZipFixtureStream() {
        var csv = File.ReadAllText(GetInputPath("sample.csv"), Encoding.UTF8);
        var markdown = "# ZIP Fixture\n\nThis archive carries markdown content.";

        var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteTextEntry(archive, "docs/readme.md", markdown);
            WriteTextEntry(archive, "data/sample.csv", csv);
        }

        ms.Position = 0;
        return ms;
    }

    private static MemoryStream BuildEpubFixtureStream() {
        var chapterText =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><h1>Fixture Chapter</h1><p>EPUB fixture body text.</p></body></html>";

        var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteTextEntry(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
            WriteTextEntry(archive, "META-INF/container.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">" +
                "<rootfiles><rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles>" +
                "</container>");

            WriteTextEntry(archive, "OEBPS/content.opf",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<package version=\"3.0\" xmlns=\"http://www.idpf.org/2007/opf\">" +
                "<manifest><item id=\"ch1\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/></manifest>" +
                "<spine><itemref idref=\"ch1\"/></spine>" +
                "</package>");

            WriteTextEntry(archive, "OEBPS/chapter.xhtml", chapterText);
        }

        ms.Position = 0;
        return ms;
    }

    private static void WriteTextEntry(ZipArchive archive, string path, string content, CompressionLevel compressionLevel = CompressionLevel.Optimal) {
        var entry = archive.CreateEntry(path, compressionLevel);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), 4096, leaveOpen: false);
        writer.Write(content);
    }
}
