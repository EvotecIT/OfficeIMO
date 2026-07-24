using System.IO.Compression;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class BrowserConversionServiceTests {
    private readonly BrowserConversionService _service = new();

    [Fact]
    public void RouteCatalog_HasUniqueCustomerRoutes() {
        Assert.Equal(6, ConversionRouteCatalog.All.Count);
        Assert.Equal(
            ConversionRouteCatalog.All.Count,
            ConversionRouteCatalog.All.Select(static route => route.Id).Distinct(StringComparer.OrdinalIgnoreCase).Count());
        Assert.All(ConversionRouteCatalog.All, static route => {
            Assert.False(string.IsNullOrWhiteSpace(route.Source));
            Assert.False(string.IsNullOrWhiteSpace(route.Target));
            Assert.False(string.IsNullOrWhiteSpace(route.EnginePath));
        });
    }

    [Fact]
    public void MarkdownToHtml_ReturnsPreviewAndDownload() {
        var route = ConversionRouteCatalog.Find("markdown-html");
        var result = _service.ConvertText(route, "# Status\n\n**Ready**");

        Assert.Equal("text/html;charset=utf-8", result.ContentType);
        Assert.Equal("officeimo-markdown.html", result.FileName);
        Assert.Contains("<h1", result.HtmlPreview, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Ready", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_UsesSharedHtmlDocument() {
        var route = ConversionRouteCatalog.Find("html-markdown");
        var result = _service.ConvertText(route, "<h1>Status</h1><p>Ready</p>");

        Assert.Equal("text/markdown;charset=utf-8", result.ContentType);
        Assert.Contains("# Status", result.Text, StringComparison.Ordinal);
        Assert.Contains("Ready", result.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownToWord_ReturnsOpenXmlPackage() {
        var route = ConversionRouteCatalog.Find("markdown-docx");
        var result = _service.ConvertText(route, "# Status\n\nReady");

        Assert.Equal("officeimo-markdown.docx", result.FileName);
        Assert.True(result.Bytes.Length > 4);
        Assert.Equal((byte)'P', result.Bytes[0]);
        Assert.Equal((byte)'K', result.Bytes[1]);
    }

    [Fact]
    public void TextConversion_RejectsInputBeyondBrowserLimit() {
        var route = ConversionRouteCatalog.Find("markdown-html");
        string oversized = new('a', BrowserConversionService.MaxTextInputChars + 1);

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() => _service.ConvertText(route, oversized));

        Assert.Contains("browser converter", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WordConversion_RejectsCompressedPackageBomb() {
        var route = ConversionRouteCatalog.Find("docx-pdf");
        byte[] bytes = CreateHighlyCompressedPackage();
        var document = new SelectedDocument("unsafe.docx", ".docx", "DOCX", bytes.LongLength, bytes);

        Assert.Throws<OfficePackageSecurityException>(() => _service.ConvertFile(route, document, limitExcelRows: false));
    }

    [Fact]
    public void WordConversion_UsesFullBrowserShapingWithoutFalseDegradationWarnings() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("مرحبا");
        byte[] bytes = source.ToBytes();
        var document = new SelectedDocument("complex-script.docx", ".docx", "DOCX", bytes.LongLength, bytes);

        ConversionResult result = _service.ConvertFile(
            ConversionRouteCatalog.Find("docx-pdf"),
            document,
            limitExcelRows: false);

        Assert.DoesNotContain(result.Warnings, warning =>
            warning.Contains("unsupported-bidirectional-text-layout", StringComparison.Ordinal));
        Assert.DoesNotContain(result.Warnings, warning =>
            warning.Contains("unsupported-complex-script-shaping", StringComparison.Ordinal));
        Assert.Equal("Faithful", result.FidelityStatus);
        Assert.NotNull(result.CompanionReport);
        Assert.Contains("NotoSansArabic", Encoding.ASCII.GetString(result.Bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void WordConversion_DoesNotInjectPageNumbers() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Faithful browser conversion");
        byte[] bytes = source.ToBytes();
        var document = new SelectedDocument("faithful.docx", ".docx", "DOCX", bytes.LongLength, bytes);

        ConversionResult result = _service.ConvertFile(
            ConversionRouteCatalog.Find("docx-pdf"),
            document,
            limitExcelRows: false);

        string text = PdfReadDocument.Open(result.Bytes).ExtractText();
        Assert.Contains("Faithful browser conversion", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Page 1", text, StringComparison.Ordinal);
        Assert.True(PdfReadDocument.Open(result.Bytes).HasTaggedContent);

        BrowserConversionArtifact report = Assert.IsType<BrowserConversionArtifact>(result.CompanionReport);
        using JsonDocument manifest = JsonDocument.Parse(report.Bytes);
        JsonElement root = manifest.RootElement;
        Assert.Equal("officeimo-browser-compact-2026.07", root.GetProperty("fontPack").GetProperty("id").GetString());
        Assert.Equal(
            "99fe9605fae25324712287bc2212236771b67515ec77dab263a35fc48079e72f",
            root.GetProperty("fontPack").GetProperty("fingerprint").GetString());
        Assert.True(root.GetProperty("output").GetProperty("tagged").GetBoolean());
        Assert.Equal(result.FidelityStatus, root.GetProperty("fidelityStatus").GetString());
        Assert.False(string.IsNullOrWhiteSpace(root.GetProperty("conversionId").GetString()));
        Assert.Empty(root.GetProperty("warnings").EnumerateArray());
        Assert.Empty(result.Warnings);
    }

    [Fact]
    public void WordConversion_PreservesBusinessDocumentStructureAndReadingOrder() {
        byte[] bytes = CreateBusinessDocument();
        var document = new SelectedDocument("business-report.docx", ".docx", "DOCX", bytes.LongLength, bytes);

        ConversionResult result = _service.ConvertFile(
            ConversionRouteCatalog.Find("docx-pdf"),
            document,
            limitExcelRows: false);

        PdfReadDocument pdf = PdfReadDocument.Open(result.Bytes);
        PdfTaggedContentInfo tagged = Assert.IsType<PdfTaggedContentInfo>(pdf.TaggedContent);
        Assert.Equal("en-US", pdf.CatalogLanguage);
        Assert.Contains("H1", tagged.StructureTypes);
        Assert.Contains("P", tagged.StructureTypes);
        Assert.Contains("L", tagged.StructureTypes);
        Assert.Contains("LI", tagged.StructureTypes);
        Assert.Contains("Table", tagged.StructureTypes);
        Assert.Contains("TH", tagged.StructureTypes);
        Assert.Contains("TD", tagged.StructureTypes);

        string text = pdf.ExtractText();
        Assert.True(text.IndexOf("Delivery status", StringComparison.Ordinal) <
                    text.IndexOf("Review owner", StringComparison.Ordinal));
        Assert.True(text.IndexOf("Review owner", StringComparison.Ordinal) <
                    text.IndexOf("Workstream", StringComparison.Ordinal));
    }

    [Fact]
    public void WordConversion_IsDeterministicForTheSameInputAndPortableProfile() {
        byte[] bytes = CreateBusinessDocument();
        var document = new SelectedDocument("deterministic.docx", ".docx", "DOCX", bytes.LongLength, bytes);
        ConversionRoute route = ConversionRouteCatalog.Find("docx-pdf");

        ConversionResult first = _service.ConvertFile(route, document, limitExcelRows: false);
        ConversionResult second = _service.ConvertFile(route, document, limitExcelRows: false);
        var renamedDocument = new SelectedDocument("renamed.docx", ".docx", "DOCX", bytes.LongLength, bytes);
        ConversionResult renamed = _service.ConvertFile(route, renamedDocument, limitExcelRows: false);

        Assert.Equal(first.Bytes, second.Bytes);
        Assert.Equal(first.FidelityStatus, second.FidelityStatus);
        Assert.Equal(first.Warnings, second.Warnings);
        using JsonDocument firstReport = JsonDocument.Parse(first.CompanionReport!.Bytes);
        using JsonDocument secondReport = JsonDocument.Parse(second.CompanionReport!.Bytes);
        Assert.Equal(
            firstReport.RootElement.GetProperty("conversionId").GetString(),
            secondReport.RootElement.GetProperty("conversionId").GetString());
        Assert.Equal(
            firstReport.RootElement.GetProperty("output").GetProperty("sha256").GetString(),
            secondReport.RootElement.GetProperty("output").GetProperty("sha256").GetString());
        using JsonDocument renamedReport = JsonDocument.Parse(renamed.CompanionReport!.Bytes);
        Assert.False(first.Bytes.SequenceEqual(renamed.Bytes));
        Assert.NotEqual(
            firstReport.RootElement.GetProperty("conversionId").GetString(),
            renamedReport.RootElement.GetProperty("conversionId").GetString());
    }

    [Fact]
    public void PdfProfiles_ExposeAccessibleAndDiagnosticContracts() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Profile contract");
        byte[] bytes = source.ToBytes();
        var document = new SelectedDocument("profile.docx", ".docx", "DOCX", bytes.LongLength, bytes);
        ConversionRoute route = ConversionRouteCatalog.Find("docx-pdf");

        ConversionResult accessible = _service.ConvertFile(
            route,
            document,
            limitExcelRows: false,
            BrowserPdfProfileCatalog.Accessible);
        ConversionResult diagnostic = _service.ConvertFile(
            route,
            document,
            limitExcelRows: false,
            BrowserPdfProfileCatalog.Diagnostic);

        Assert.Equal("accessible", accessible.Profile?.Id);
        Assert.Contains("pdfuaid:part", Encoding.UTF8.GetString(accessible.Bytes), StringComparison.Ordinal);
        BrowserConversionArtifact overlay = Assert.IsType<BrowserConversionArtifact>(diagnostic.DebugOverlay);
        Assert.Equal("profile.page-1.layout.svg", overlay.FileName);
        Assert.Contains("<svg", Encoding.UTF8.GetString(overlay.Bytes), StringComparison.Ordinal);
        using JsonDocument manifest = JsonDocument.Parse(diagnostic.CompanionReport!.Bytes);
        Assert.Equal(
            "diagnostic",
            manifest.RootElement.GetProperty("engine").GetProperty("profile").GetProperty("id").GetString());
        Assert.True(manifest.RootElement.GetProperty("performance").GetProperty("peakRetainedCompletedPayloadBytes").GetInt64() > 0);
        Assert.True(manifest.RootElement.GetProperty("performance").GetProperty("isForwardOnlyObjectSerialization").GetBoolean());
        Assert.False(manifest.RootElement.GetProperty("performance").GetProperty("isForwardOnlyLayout").GetBoolean());
    }

    [Fact]
    public void AccessiblePdfProfile_PreservesTheWordSourceLanguage() {
        using WordDocument source = WordDocument.Create();
        source.Settings.Language = "pl-PL";
        source.AddParagraph("Dokument dostępny");
        byte[] bytes = source.ToBytes();
        var document = new SelectedDocument("dostepny.docx", ".docx", "DOCX", bytes.LongLength, bytes);

        ConversionResult accessible = _service.ConvertFile(
            ConversionRouteCatalog.Find("docx-pdf"),
            document,
            limitExcelRows: false,
            BrowserPdfProfileCatalog.Accessible);

        Assert.Equal("pl-PL", PdfReadDocument.Open(accessible.Bytes).CatalogLanguage);
    }

    [Fact]
    public void SupportBundle_ExcludesDocumentContentUnlessExplicitlyIncluded() {
        using WordDocument source = WordDocument.Create();
        source.AddParagraph("Private customer content marker");
        byte[] bytes = source.ToBytes();
        var document = new SelectedDocument("private-name.docx", ".docx", "DOCX", bytes.LongLength, bytes);
        ConversionResult result = _service.ConvertFile(
            ConversionRouteCatalog.Find("docx-pdf"),
            document,
            limitExcelRows: false);

        BrowserConversionArtifact safe = _service.CreateSupportBundle(document, result);
        using var safeStream = new MemoryStream(safe.Bytes);
        using var safeArchive = new ZipArchive(safeStream, ZipArchiveMode.Read);
        Assert.Equal(["README.txt", "support-summary.json"], safeArchive.Entries.Select(static entry => entry.FullName).ToArray());
        Assert.DoesNotContain("private-name", Encoding.UTF8.GetString(safe.Bytes), StringComparison.Ordinal);
        Assert.DoesNotContain("Private customer content marker", Encoding.UTF8.GetString(safe.Bytes), StringComparison.Ordinal);

        BrowserConversionArtifact explicitContent = _service.CreateSupportBundle(
            document,
            result,
            includeDocumentContent: true);
        using var contentStream = new MemoryStream(explicitContent.Bytes);
        using var contentArchive = new ZipArchive(contentStream, ZipArchiveMode.Read);
        Assert.Contains(contentArchive.Entries, static entry => entry.FullName == "content/source.docx");
        Assert.Contains(contentArchive.Entries, static entry => entry.FullName == "content/result.pdf");
    }

    [Fact]
    public void ExcelConversion_UsesThePortableTaggedPdfProfile() {
        using ExcelDocument workbook = ExcelDocument.Create();
        var sheet = workbook.AddWorksheet("Delivery");
        sheet.CellValue(1, 1, "Workstream");
        sheet.CellValue(1, 2, "Status");
        sheet.CellValue(2, 1, "Platform");
        sheet.CellValue(2, 2, "Ready");
        byte[] bytes = workbook.ToBytes();
        var document = new SelectedDocument("delivery.xlsx", ".xlsx", "XLSX", bytes.LongLength, bytes);

        ConversionResult result = _service.ConvertFile(
            ConversionRouteCatalog.Find("xlsx-pdf"),
            document,
            limitExcelRows: false);

        PdfReadDocument pdf = PdfReadDocument.Open(result.Bytes);
        Assert.True(pdf.HasTaggedContent);
        Assert.Contains("Workstream", pdf.ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Platform", pdf.ExtractText(), StringComparison.Ordinal);
        BrowserConversionArtifact report = Assert.IsType<BrowserConversionArtifact>(result.CompanionReport);
        Assert.Equal("delivery.conversion.json", report.FileName);
        Assert.Contains("OfficeIMO.Excel.Pdf", Encoding.UTF8.GetString(report.Bytes), StringComparison.Ordinal);

        ConversionResult limitedResult = _service.ConvertFile(
            ConversionRouteCatalog.Find("xlsx-pdf"),
            document,
            limitExcelRows: true);
        using JsonDocument unlimitedManifest = JsonDocument.Parse(report.Bytes);
        using JsonDocument limitedManifest = JsonDocument.Parse(limitedResult.CompanionReport!.Bytes);
        Assert.NotEqual(
            unlimitedManifest.RootElement.GetProperty("conversionId").GetString(),
            limitedManifest.RootElement.GetProperty("conversionId").GetString());
        Assert.Equal(
            "maxRowsPerSheet=unlimited",
            unlimitedManifest.RootElement.GetProperty("engine").GetProperty("optionProfile").GetString());
        Assert.Equal(
            "maxRowsPerSheet=250",
            limitedManifest.RootElement.GetProperty("engine").GetProperty("optionProfile").GetString());
    }

    [Fact]
    public void PowerPointConversion_UsesThePortableTaggedPdfProfile() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.AddSlide().AddTextBoxPoints("Delivery readiness", 36, 36, 320, 64);
        byte[] bytes = presentation.ToBytes();
        var document = new SelectedDocument("readiness.pptx", ".pptx", "PPTX", bytes.LongLength, bytes);

        ConversionResult result = _service.ConvertFile(
            ConversionRouteCatalog.Find("pptx-pdf"),
            document,
            limitExcelRows: false);

        PdfReadDocument pdf = PdfReadDocument.Open(result.Bytes);
        Assert.True(pdf.HasTaggedContent);
        Assert.Contains("Delivery readiness", pdf.ExtractText(), StringComparison.Ordinal);
        BrowserConversionArtifact report = Assert.IsType<BrowserConversionArtifact>(result.CompanionReport);
        Assert.Equal("readiness.conversion.json", report.FileName);
        Assert.Contains("OfficeIMO.PowerPoint.Pdf", Encoding.UTF8.GetString(report.Bytes), StringComparison.Ordinal);
    }

    private static byte[] CreateBusinessDocument() {
        using WordDocument source = WordDocument.Create();
        WordParagraph heading = source.AddParagraph("Delivery status");
        heading.Style = WordParagraphStyles.Heading1;
        source.AddParagraph("The current review is ready for customer acceptance.");
        WordList list = source.AddList(WordListStyle.Bulleted);
        list.AddItem("Review owner");
        list.AddItem("Acceptance criteria");
        WordTable table = source.AddTable(2, 2, WordTableStyle.TableGrid);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Workstream";
        table.Rows[0].Cells[1].Paragraphs[0].Text = "Status";
        table.Rows[1].Cells[0].Paragraphs[0].Text = "Platform";
        table.Rows[1].Cells[1].Paragraphs[0].Text = "Ready";
        return source.ToBytes();
    }

    private static byte[] CreateHighlyCompressedPackage() {
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true)) {
            ZipArchiveEntry contentTypes = archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);
            using (var writer = new StreamWriter(contentTypes.Open())) {
                writer.Write("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\" />");
            }

            ZipArchiveEntry oversizedPart = archive.CreateEntry("word/document.xml", CompressionLevel.Optimal);
            using var stream = oversizedPart.Open();
            byte[] repeated = new byte[2 * 1024 * 1024];
            Array.Fill(repeated, (byte)'A');
            stream.Write(repeated);
        }
        return buffer.ToArray();
    }
}
