using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderPdfManagedInspectionEvidenceTests {
    [Fact]
    public void TextAndMetadataRetainsMappingFailuresAndTerminalEvidenceLimits() {
        var reader = new OfficeDocumentReaderBuilder().AddPdfHandler().Build();
        var options = new ReaderOptions { Projection = ReaderProjection.TextAndMetadata, Inspection = new ReaderInspectionOptions() };
        OfficeDocumentReadResult result = reader.ReadDocument(BuildMixedPdf(), "mixed.pdf", options);
        Assert.Null(result.Markdown); Assert.Empty(result.Assets); Assert.Empty(result.OcrCandidates);
        Assert.Contains(result.Chunks, chunk => chunk.Location.Page == 2 && chunk.Warnings?.Count > 0);
        Assert.Contains(result.Metadata, entry => entry.Category == "pdf.page-outcome" && entry.Value == "UnmappableText");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "pdf-page-unmappable-text");
        options.Inspection = new ReaderInspectionOptions { MaxRecords = 1 };
        Assert.Throws<ReaderResourceLimitException>(() => reader.ReadDocument(BuildMixedPdf(), "mixed.pdf", options));
    }

    [Fact]
    public void OptInInspectionRetainsReadableTextFailedPageStructureAndV14Evidence() {
        byte[] bytes = BuildMixedPdf();
        var reader = new OfficeDocumentReaderBuilder().AddPdfHandler().Build();
        Assert.Throws<PdfUnsupportedTextMappingException>(() => reader.ReadDocument(bytes, "mixed.pdf"));
        var options = new ReaderOptions { Inspection = new ReaderInspectionOptions() };
        OfficeDocumentReadResult result = reader.ReadDocument(bytes, "mixed.pdf", options);
        Assert.Single(result.Pages);
        Assert.Contains("Readable native text", result.Markdown);
        Assert.Contains(result.Chunks, chunk => chunk.Location.Page == 2 && chunk.Warnings?.Count > 0);
        Assert.Contains(result.Metadata, entry => entry.Category == "pdf.page-outcome" && entry.Location?.Page == 2 && entry.Value == "UnmappableText");
        Assert.Contains(result.Metadata, entry => entry.Category == "pdf.source.link" && entry.Location?.Page == 2 && entry.Value == "https://example.test/failed-page");
        Assert.Contains(result.Metadata, entry => entry.Category == "pdf.text-decoding" && entry.Location?.Page == 1 && entry.Attributes["source"] == "FontEncoding");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "pdf-page-unmappable-text" && diagnostic.Location?.Page == 2);
        string json = OfficeDocumentReadResultJson.Serialize(result);
        using JsonDocument parsed = JsonDocument.Parse(json);
        Assert.Equal(14, result.SchemaVersion);
        Assert.Equal(14, parsed.RootElement.GetProperty("schemaVersion").GetInt32());
        OfficeDocumentReadResult roundTrip = OfficeDocumentReadResultJson.Deserialize(json);
        Assert.Contains(roundTrip.Metadata, entry => entry.Category == "pdf.page-outcome" && entry.Value == "UnmappableText");
        OfficeDocumentReadResult fallback = reader.ReadDocument(BuildMixedPdf(fallbackFont: true), "fallback.pdf", options);
        Assert.Contains(fallback.Diagnostics, diagnostic => diagnostic.Code == "pdf-text-decoding-fallback" && diagnostic.Location?.Page == 1);
    }

    [Fact]
    public void FullParserBudgetsAndSharedEvidenceBudgetsAreTerminalThroughThePublicReader() {
        byte[] bytes = BuildMixedPdf();
        var reader = new OfficeDocumentReaderBuilder().AddPdfHandler(new ReaderPdfOptions {
            LoadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxPages = 1 } }
        }).Build();
        var inspection = new ReaderOptions { Inspection = new ReaderInspectionOptions() };
        Assert.Throws<PdfReadLimitException>(() => reader.ReadDocument(bytes, "mixed.pdf", inspection));
        var ordinary = new OfficeDocumentReaderBuilder().AddPdfHandler().Build();
        Assert.Throws<ReaderResourceLimitException>(() => ordinary.ReadDocument(bytes, "mixed.pdf",
            new ReaderOptions { Inspection = new ReaderInspectionOptions { MaxRecords = 1 } }));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => ordinary.ReadDocument(bytes, "mixed.pdf", inspection, cancellation.Token));
    }

    [Fact]
    public void RegistrationSnapshotsParserOptionsAndReaderInputLimitFurtherRestrictsThem() {
        byte[] bytes = BuildMixedPdf();
        var configured = new ReaderPdfOptions {
            LoadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxPages = 3, MaxInputBytes = bytes.Length + 1 } }
        };
        var reader = new OfficeDocumentReaderBuilder().AddPdfHandler(configured).Build();
        configured.LoadOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxPages = 1 } };
        var inspection = new ReaderOptions { Inspection = new ReaderInspectionOptions() };
        Assert.Single(reader.ReadDocument(bytes, "mixed.pdf", inspection).Pages);
        Assert.ThrowsAny<IOException>(() => reader.ReadDocument(bytes, "mixed.pdf",
            new ReaderOptions { MaxInputBytes = bytes.Length - 1, Inspection = new ReaderInspectionOptions() }));
        ReaderPdfOptions clone = new ReaderPdfOptions {
            LoadOptions = new PdfLoadOptions { ParsingMode = PdfParsingMode.Strict, Limits = new PdfReadLimits { MaxJavaScriptBytes = 713, MaxContentOperations = 241 } }
        }.Clone();
        Assert.Equal(PdfParsingMode.Strict, clone.LoadOptions!.ParsingMode);
        Assert.Equal(713, clone.LoadOptions.Limits.MaxJavaScriptBytes);
        Assert.Equal(241, clone.LoadOptions.Limits.MaxContentOperations);
    }

    private static byte[] BuildMixedPdf(bool fallbackFont = false) {
        string[] objects = {
            "<< /Type /Catalog /Pages 2 0 R >>", "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 5 0 R >> >> /Contents 7 0 R >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 6 0 R >> >> /Contents 8 0 R /Annots [9 0 R] >>",
            fallbackFont ? "<< /Type /Font /Subtype /Type1 /BaseFont /UnknownFont >>"
                : "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "<< /Type /Font /Subtype /Type0 /BaseFont /Unmapped /Encoding /Mystery-H >>",
            Stream("BT /F1 12 Tf 20 100 Td (Readable native text) Tj ET"),
            Stream("BT /F1 12 Tf 20 100 Td <0001> Tj ET"),
            "<< /Type /Annot /Subtype /Link /Rect [20 30 80 50] /A << /S /URI /URI (https://example.test/failed-page) >> >>"
        };
        var text = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>();
        for (int index = 0; index < objects.Length; index++) {
            offsets.Add(text.Length);
            text.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = text.Length;
        text.Append("xref\n0 ").Append(objects.Length + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) text.Append(offset.ToString("D10", CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        text.Append("trailer\n<< /Root 1 0 R /Size ").Append(objects.Length + 1).Append(" >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(text.ToString());
    }

    private static string Stream(string content) => "<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) +
        " >>\nstream\n" + content + "\nendstream";
}
