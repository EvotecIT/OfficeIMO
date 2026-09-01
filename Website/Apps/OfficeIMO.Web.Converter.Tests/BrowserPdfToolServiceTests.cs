using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class BrowserPdfToolServiceTests {
    private readonly BrowserPdfToolService _service = new();

    [Fact]
    public void Catalog_ExposesUniqueTaskOrientedTools() {
        Assert.Equal(12, PdfToolCatalog.All.Count);
        Assert.Equal(12, PdfToolCatalog.All.Select(static tool => tool.Id).Distinct(StringComparer.OrdinalIgnoreCase).Count());
        Assert.Contains(PdfToolCatalog.All, static tool => tool.Kind == PdfToolKind.Redact && tool.RequiresDestructiveConfirmation);
        Assert.Contains(PdfToolCatalog.All, static tool => tool.Kind == PdfToolKind.Compare && tool.InputMode == PdfToolInputMode.Pair);
    }

    [Fact]
    public void Inspect_ReturnsBoundedMachineReadablePreflight() {
        PdfToolResult result = _service.Execute(Request("inspect", [Document(CreatePdf("Inspection evidence"))]));

        Assert.Equal("application/json", result.Artifact.ContentType);
        using JsonDocument report = JsonDocument.Parse(result.Artifact.Bytes);
        Assert.Equal("inspect", report.RootElement.GetProperty("tool").GetString());
        Assert.Equal("OfficeIMO.Pdf", report.RootElement.GetProperty("engine").GetString());
        Assert.True(report.RootElement.GetProperty("browserLocal").GetBoolean());
        Assert.Equal("True", report.RootElement.GetProperty("details").GetProperty("canRead").GetString());
        Assert.Equal(1, result.PageCount);
        Assert.NotNull(result.Report);
        using JsonDocument operationReport = JsonDocument.Parse(result.Report!.Bytes);
        JsonElement output = operationReport.RootElement.GetProperty("output");
        Assert.Equal(result.Artifact.Bytes.LongLength, output.GetProperty("bytes").GetInt64());
        Assert.Equal(
            Convert.ToHexString(SHA256.HashData(result.Artifact.Bytes)).ToLowerInvariant(),
            output.GetProperty("sha256").GetString());
    }

    [Fact]
    public void MergeAndSplit_ProduceReadablePdfArtifacts() {
        SelectedDocument first = Document(CreatePdf("First"), "first.pdf");
        SelectedDocument second = Document(CreatePdf("Second"), "second.pdf");
        PdfToolResult merged = _service.Execute(Request("merge", [first, second]));

        Assert.Equal(2, PdfDocument.Load(merged.Artifact.Bytes).Inspect().PageCount);
        Assert.NotNull(merged.Report);

        PdfToolResult split = _service.Execute(Request("split", [Document(CreateThreePagePdf())], pagesPerDocument: 2));
        using var archive = new ZipArchive(new MemoryStream(split.Artifact.Bytes), ZipArchiveMode.Read);
        Assert.Equal(2, archive.Entries.Count);
        Assert.All(archive.Entries, static entry => Assert.EndsWith(".pdf", entry.Name, StringComparison.OrdinalIgnoreCase));

        InvalidDataException tooManyParts = Assert.Throws<InvalidDataException>(() =>
            _service.Execute(Request("split", [Document(CreatePdfWithPages(BrowserPdfPolicy.MaxSplitDocuments + 1))])));
        Assert.Contains("browser limit", tooManyParts.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PageTools_ApplySelectorsAndRequireDestructiveConfirmation() {
        SelectedDocument source = Document(CreateThreePagePdf());
        PdfToolResult extracted = _service.Execute(Request("extract", [source], pageSelection: "3..2"));
        Assert.Equal(2, PdfDocument.Load(extracted.Artifact.Bytes).Inspect().PageCount);

        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() =>
            _service.Execute(Request("delete", [source], pageSelection: "1")));
        Assert.Contains("Confirm", error.Message, StringComparison.Ordinal);

        PdfToolResult deleted = _service.Execute(Request("delete", [source], pageSelection: "1", confirmed: true));
        Assert.Equal(2, PdfDocument.Load(deleted.Artifact.Bytes).Inspect().PageCount);
    }

    [Fact]
    public void Optimize_StaysLosslessAndReportsItsDecision() {
        PdfToolResult result = _service.Execute(Request("optimize", [Document(CreatePdf("Lossless"))]));

        Assert.Equal("application/pdf", result.Artifact.ContentType);
        Assert.Equal(1, PdfDocument.Load(result.Artifact.Bytes).Inspect().PageCount);
        using JsonDocument report = JsonDocument.Parse(result.Report!.Bytes);
        Assert.Equal("Balanced", report.RootElement.GetProperty("details").GetProperty("profile").GetString());
    }

    [Fact]
    public void ProtectAndUnlock_RoundTripStandardEncryption() {
        PdfToolResult protectedPdf = _service.Execute(Request(
            "protect",
            [Document(CreatePdf("Protected"))],
            userPassword: "reader-2026",
            ownerPassword: "owner-2026"));
        Assert.True(PdfDocument.Preflight(protectedPdf.Artifact.Bytes).Probe.Security.HasEncryption);
        using (JsonDocument protectionReport = JsonDocument.Parse(protectedPdf.Report!.Bytes)) {
            JsonElement details = protectionReport.RootElement.GetProperty("details");
            Assert.Equal("True", details.GetProperty("preservationVerified").GetString());
            Assert.Equal("0", details.GetProperty("preservationIssueCount").GetString());
            Assert.Contains("passed", details.GetProperty("preservationSummary").GetString(), StringComparison.OrdinalIgnoreCase);
        }

        PdfToolResult unlocked = _service.Execute(Request(
            "unlock",
            [Document(protectedPdf.Artifact.Bytes, protectedPdf.Artifact.FileName)],
            ownerPassword: "owner-2026"));
        Assert.False(PdfDocument.Preflight(unlocked.Artifact.Bytes).Probe.Security.HasEncryption);
        Assert.Contains("Protected", PdfReadDocument.Open(unlocked.Artifact.Bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void ProtectAndUnlock_RoundTripJapaneseArchivalArtifact() {
        const string japanese = "日本語の保存版を東京都で確認します。";
        var conversions = new BrowserConversionService();
        ConversionResult archival = conversions.ConvertText(
            ConversionRouteCatalog.Find("html-pdf"),
            $"<article lang='ja'><h1>日本語の保存版</h1><p>{japanese}</p></article>",
            BrowserPdfProfileCatalog.Archival);

        PdfToolResult protectedPdf = _service.Execute(Request(
            "protect",
            [Document(archival.Bytes, archival.FileName)],
            userPassword: "reader-2026",
            ownerPassword: "owner-2026"));
        PdfToolResult unlocked = _service.Execute(Request(
            "unlock",
            [Document(protectedPdf.Artifact.Bytes, protectedPdf.Artifact.FileName)],
            ownerPassword: "owner-2026"));

        PdfDocumentPreflight protectedPreflight = PdfDocument.Preflight(protectedPdf.Artifact.Bytes, new PdfLoadOptions {
            Password = "reader-2026",
            AesCryptographyProvider = OfficeManagedAesCryptographyProvider.Default
        });
        Assert.True(protectedPreflight.Probe.Security.HasEncryption);
        Assert.Equal(6, protectedPreflight.Probe.Security.EncryptionRevision);
        Assert.Equal(256, protectedPreflight.Probe.Security.EncryptionLengthBits);
        Assert.False(PdfDocument.Preflight(unlocked.Artifact.Bytes).Probe.Security.HasEncryption);
        Assert.Contains(japanese, PdfReadDocument.Open(unlocked.Artifact.Bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void Redact_RemovesAndVerifiesLiteralMarker() {
        const string marker = "PAY-SECRET-2026";
        PdfToolResult result = _service.Execute(Request(
            "redact",
            [Document(CreatePdf("Public " + marker + " public"))],
            redactionText: marker,
            confirmed: true));

        Assert.DoesNotContain(marker, PdfReadDocument.Open(result.Artifact.Bytes).ExtractText(), StringComparison.OrdinalIgnoreCase);
        using JsonDocument report = JsonDocument.Parse(result.Report!.Bytes);
        Assert.Equal("True", report.RootElement.GetProperty("details").GetProperty("verified").GetString());
    }

    [Fact]
    public void Redact_CaseInsensitiveSearchRejectsConcreteCaseVariantResidue() {
        byte[] source = PdfDocument.Create(pdf => pdf.Content(content => content
                .Paragraph(paragraph => paragraph.Text("Public secret public"))))
            .Meta(title: "SECRET")
            .ToBytes();

        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => _service.Execute(Request(
            "redact",
            [Document(source)],
            redactionText: "secret",
            confirmed: true)));

        Assert.Contains("verification", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void BrowserReadPolicy_UsesWebAssemblyScaleBudgets() {
        PdfLoadOptions options = BrowserPdfPolicy.CreateReadOptions();

        Assert.Same(OfficeManagedAesCryptographyProvider.Default, options.AesCryptographyProvider);
        Assert.Equal(BrowserPdfPolicy.MaxInputBytes, options.Limits.MaxInputBytes);
        Assert.Equal(BrowserPdfPolicy.MaxPages, options.Limits.MaxPages);
        Assert.True(options.Limits.MaxDecodedStreamBytes <= 32 * 1024 * 1024);
        Assert.True(options.Limits.MaxTotalDecodedStreamBytes <= 96L * 1024L * 1024L);
        Assert.True(options.Limits.MaxIndirectObjects <= 50_000);
    }

    [Fact]
    public void Compare_IdenticalPdfReturnsSelfContainedMatchGallery() {
        byte[] bytes = CreatePdf("Same visual");
        PdfToolResult result = _service.Execute(Request(
            "compare",
            [Document(bytes, "expected.pdf"), Document((byte[])bytes.Clone(), "actual.pdf")]));

        Assert.StartsWith("text/html", result.Artifact.ContentType, StringComparison.Ordinal);
        Assert.Contains("<!doctype html", Encoding.UTF8.GetString(result.Artifact.Bytes), StringComparison.OrdinalIgnoreCase);
        using JsonDocument report = JsonDocument.Parse(result.Report!.Bytes);
        Assert.Equal("True", report.RootElement.GetProperty("details").GetProperty("isMatch").GetString());
    }

    private static PdfToolRequest Request(
        string tool,
        IReadOnlyList<SelectedDocument> files,
        string pageSelection = "all",
        int pagesPerDocument = 1,
        string userPassword = "",
        string ownerPassword = "",
        string redactionText = "",
        bool confirmed = false) =>
        new(
            PdfToolCatalog.Find(tool),
            files,
            pageSelection,
            pagesPerDocument,
            90,
            PdfOptimizationProfile.Balanced,
            userPassword,
            ownerPassword,
            redactionText,
            confirmed);

    private static SelectedDocument Document(byte[] bytes, string name = "source.pdf") =>
        new(name, ".pdf", "PDF", bytes.LongLength, bytes);

    private static byte[] CreatePdf(string text) =>
        PdfDocument.Create(pdf => pdf.Content(content => content.Paragraph(paragraph => paragraph.Text(text)))).ToBytes();

    private static byte[] CreateThreePagePdf() =>
        PdfDocument.Create(pdf => pdf.Content(content => content
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Page two"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Page three"))))
            .ToBytes();

    private static byte[] CreatePdfWithPages(int pageCount) {
        if (pageCount <= 0) throw new ArgumentOutOfRangeException(nameof(pageCount));
        return PdfDocument.Create(pdf => pdf.Content(content => {
            for (int page = 1; page <= pageCount; page++) {
                content.Paragraph(paragraph => paragraph.Text($"Page {page}"));
                if (page < pageCount) content.PageBreak();
            }
        })).ToBytes();
    }
}
