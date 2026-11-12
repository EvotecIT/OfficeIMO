using System.IO.Compression;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Tool.Agent;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tool.Tests;

// The CLI resolves these routes from the shared catalog; each family proves a real artifact through `officeimo convert`.
public sealed class HostConversionRouteTests {
    private const string Marker = "Host route marker";

    [Theory]
    [InlineData(".docx", ".odt")]
    [InlineData(".xlsx", ".ods")]
    [InlineData(".pptx", ".odp")]
    public async Task OpenDocumentRoutesConvertToPdfAndBackToOffice(string office, string openDocument) {
        string directory = CreateDirectory();
        try {
            string source = Path.Combine(directory, "source" + office);
            string odf = Path.Combine(directory, "source" + openDocument);
            string pdf = Path.Combine(directory, "source.pdf");
            string back = Path.Combine(directory, "back" + office);
            string markdown = Path.Combine(directory, "back.md");
            CreateOffice(source);

            await ConvertAsync(source, odf);
            await ConvertAsync(odf, pdf);
            await ConvertAsync(odf, back);
            await ConvertAsync(back, markdown);

            AssertPdfText(pdf, Marker);
            Assert.Contains(Marker, await File.ReadAllTextAsync(markdown), StringComparison.Ordinal);
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(".md", "# Host route marker\n\nMarkdown body.")]
    [InlineData(".html", "<html><body><h1>Host route marker</h1><p>HTML body.</p></body></html>")]
    [InlineData(".rtf", @"{\rtf1\ansi Host route marker\par RTF body.\par}")]
    public async Task MarkupRoutesWriteEditableWordDocuments(string extension, string content) {
        string directory = CreateDirectory();
        try {
            string source = Path.Combine(directory, "source" + extension);
            string docx = Path.Combine(directory, "result.docx");
            await File.WriteAllTextAsync(source, content);

            await ConvertAsync(source, docx);

            using WordDocument document = WordDocument.Load(docx);
            Assert.Contains(document.Paragraphs, paragraph => paragraph.Text.Contains(Marker, StringComparison.Ordinal));
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task EmailRouteRendersTheStyledMessageAndTextOnlyKeepsTheReaderProjection() {
        string directory = CreateDirectory();
        try {
            string source = Path.Combine(directory, "message.eml");
            string styled = Path.Combine(directory, "styled.pdf");
            string archival = Path.Combine(directory, "archival.pdf");
            await File.WriteAllTextAsync(source,
                "From: sender@example.test\r\nTo: recipient@example.test\r\nSubject: Host route marker\r\n" +
                "MIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary=\"part\"\r\n\r\n" +
                "--part\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nMessage body.\r\n" +
                "--part\r\nContent-Type: text/plain; name=\"notes.txt\"\r\nContent-Disposition: attachment; filename=\"notes.txt\"\r\n\r\n" +
                "attached\r\n--part--\r\n");

            await ConvertAsync(source, styled);
            await ConvertAsync(source, archival, "--text-only");

            string rendered = AssertPdfText(styled, Marker);
            Assert.Contains("Attachments (1)", rendered, StringComparison.Ordinal);
            Assert.Contains("notes.txt", rendered, StringComparison.Ordinal);
            Assert.DoesNotContain("Attachments (1)", AssertPdfText(archival, "Message body."), StringComparison.Ordinal);
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task EpubRouteWritesChaptersInSpineOrder() {
        string directory = CreateDirectory();
        try {
            string source = Path.Combine(directory, "book.epub");
            string pdf = Path.Combine(directory, "book.pdf");
            CreateEpub(source);

            await ConvertAsync(source, pdf);

            string text = AssertPdfText(pdf, "First chapter");
            Assert.True(text.IndexOf("First chapter", StringComparison.Ordinal) < text.IndexOf("Second chapter", StringComparison.Ordinal));
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task VisioRouteWritesTheDiagramPdf() {
        string directory = CreateDirectory();
        try {
            string source = Path.Combine(directory, "diagram.vsdx");
            string pdf = Path.Combine(directory, "diagram.pdf");
            VisioDocument diagram = VisioDocument.Create();
            VisioPage page = diagram.AddPage("Topology", OfficeLength.Inches(8), OfficeLength.Inches(5));
            page.Shapes.Add(new VisioShape("gateway") { Text = Marker });
            await File.WriteAllBytesAsync(source, diagram.ToBytes());

            await ConvertAsync(source, pdf);

            AssertPdfText(pdf, Marker);
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task DataRoutesWriteConvertedDataWithTheSelectedRowsAndColumns() {
        string directory = CreateDirectory();
        try {
            string yaml = Path.Combine(directory, "response.yaml");
            string json = Path.Combine(directory, "response.json");
            string csv = Path.Combine(directory, "rows.csv");
            await File.WriteAllTextAsync(yaml, "data:\n  items:\n    - id: 1\n      name: first\n    - id: 2\n      name: second\n");

            await ConvertAsync(yaml, json);
            await ConvertAsync(json, csv, "--select", "data.items", "--columns", "name,id");

            using (System.Text.Json.JsonDocument data = System.Text.Json.JsonDocument.Parse(await File.ReadAllTextAsync(json)))
                Assert.Equal(2, data.RootElement.GetProperty("data").GetProperty("items")[1].GetProperty("id").GetInt32());
            Assert.Equal(["name,id", "first,1", "second,2"], (await File.ReadAllTextAsync(csv)).Split('\n', StringSplitOptions.RemoveEmptyEntries));
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task DataRoutesReportLossAndRequireNoLossRefusesToPublish() {
        string directory = CreateDirectory();
        try {
            string source = Path.Combine(directory, "settings.json");
            string target = Path.Combine(directory, "settings.toml");
            await File.WriteAllTextAsync(source, "{\"name\":\"a\",\"note\":null}");
            await using var output = new MemoryStream();
            using var error = new StringWriter();

            int exit = await OfficeImoToolApp.RunAsync(["convert", source, target, "--require-no-loss"], Stream.Null, output, error);

            Assert.Equal((int)OfficeImoToolExitCode.OperationFailed, exit);
            Assert.Contains("null-dropped", error.ToString(), StringComparison.Ordinal);
            Assert.False(File.Exists(target));
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("source.docx", "result.ods", null)]
    [InlineData("notes.txt", "notes.docx", null)]
    [InlineData("source.docx", "result.pdf", "--text-only")]
    public async Task ConvertRejectsPairsWithoutOneExecutableRoute(string input, string destination, string? option) {
        await using var output = new MemoryStream();
        using var error = new StringWriter();
        string[] args = option is null ? ["convert", input, destination] : ["convert", input, destination, option];

        int exit = await OfficeImoToolApp.RunAsync(args, Stream.Null, output, error);

        Assert.Equal((int)OfficeImoToolExitCode.Usage, exit);
        Assert.Empty(output.ToArray());
    }

    [Fact]
    public async Task McpConvertRunsCatalogRoutesWithinTheAllowedRoots() {
        string directory = CreateDirectory();
        try {
            string source = Path.Combine(directory, "source.docx");
            string destination = Path.Combine(directory, "result.odt");
            CreateOffice(source);
            var service = new OfficeImoAgentService();

            AgentConvertResult result = await service.ConvertAsync(source, destination, "odt");

            Assert.True(result.Workflow?.Succeeded, result.Workflow?.Summary);
            Assert.Equal("odt", result.Format);
            Assert.Equal(64, result.Sha256.Length);
            Assert.True(File.Exists(destination));
            AgentUsageException rejected = await Assert.ThrowsAsync<AgentUsageException>(
                () => service.ConvertAsync(source, Path.Combine(directory, "result.ods"), "ods"));
            Assert.Contains("No conversion route", rejected.Message, StringComparison.Ordinal);
        } finally { Directory.Delete(directory, true); }
    }

    private static async Task ConvertAsync(string input, string output, params string[] options) {
        await using var standardOutput = new MemoryStream();
        using var error = new StringWriter();
        int exit = await OfficeImoToolApp.RunAsync(["convert", input, output, .. options], Stream.Null, standardOutput, error);
        Assert.True(exit == (int)OfficeImoToolExitCode.Success, error.ToString());
    }

    private static string AssertPdfText(string path, string expected) {
        using PdfDocument document = PdfDocument.Load(path);
        // Layout may wrap a short label, so compare whitespace-normalized text.
        string text = string.Join(" ", document.ExtractText().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        Assert.Contains(expected, text, StringComparison.Ordinal);
        return text;
    }

    private static void CreateOffice(string path) {
        switch (Path.GetExtension(path)) {
            case ".docx":
                using (WordDocument document = WordDocument.Create(path)) {
                    document.AddParagraph(Marker);
                    document.Save();
                }
                break;
            case ".xlsx":
                using (ExcelDocument workbook = ExcelDocument.Create(path)) {
                    workbook.AddSheet("Summary").Cell(1, 1).SetValue(Marker);
                    workbook.Save();
                }
                break;
            case ".pptx":
                using (PowerPointDocument presentation = PowerPointDocument.Create(path)) {
                    presentation.AddSlide().AddTextBox(Marker);
                    presentation.Save();
                }
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(path));
        }
    }

    private static void CreateEpub(string path) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Write(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
        Write(archive, "META-INF/container.xml",
            "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\"><rootfiles>" +
            "<rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles></container>");
        Write(archive, "OEBPS/content.opf",
            "<package version=\"3.0\" xmlns=\"http://www.idpf.org/2007/opf\"><metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>Host book</dc:title></metadata><manifest>" +
            "<item id=\"one\" href=\"one.xhtml\" media-type=\"application/xhtml+xml\"/><item id=\"two\" href=\"two.xhtml\" media-type=\"application/xhtml+xml\"/>" +
            "</manifest><spine><itemref idref=\"one\"/><itemref idref=\"two\"/></spine></package>");
        Write(archive, "OEBPS/one.xhtml", "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><h1>First chapter</h1></body></html>");
        Write(archive, "OEBPS/two.xhtml", "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><h1>Second chapter</h1></body></html>");
    }

    private static void Write(ZipArchive archive, string name, string content, CompressionLevel level = CompressionLevel.Optimal) {
        using var writer = new StreamWriter(archive.CreateEntry(name, level).Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private static string CreateDirectory() =>
        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "OfficeIMO.Tool.Tests", Guid.NewGuid().ToString("N"))).FullName;
}
