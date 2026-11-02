using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderPdfManagedInspectionEvidenceTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void FailedSelectionsDoNotRenumberLaterRepeatedReadablePages(bool lean, bool chunksOnly) {
        var reader = new OfficeDocumentReaderBuilder().AddPdfHandler(new ReaderPdfOptions {
            ReadOptions = new PdfLoadOptions { PageSelection = PdfPageSelection.From(2, 1, 1) }
        }).Build();
        var options = new ReaderOptions { Inspection = new ReaderInspectionOptions(),
            Projection = lean ? ReaderProjection.TextAndMetadata : ReaderProjection.Full };
        byte[] bytes = BuildMixedPdf(decoratedReadablePage: true);
        OfficeDocumentReadResult? result = chunksOnly ? null : reader.ReadDocument(bytes, "selected.pdf", options);
        ReaderChunk[] chunks = (chunksOnly ? reader.Read(bytes, "selected.pdf", options) : result!.Chunks)
            .Where(chunk => chunk.Warnings == null || chunk.Warnings.Count == 0).ToArray();
        Assert.Equal(new int?[] { 1, 2 }, chunks.Select(chunk => chunk.Location.SourceBlockIndex));
        Assert.EndsWith("selection-0001", chunks[0].Location.BlockAnchor);
        Assert.EndsWith("selection-0002", chunks[1].Location.BlockAnchor);
        Assert.NotEqual(chunks[0].Id, chunks[1].Id);
        if (!lean) Assert.Equal(new int?[] { 1, 2 }, chunks.SelectMany(chunk => chunk.Visuals!).Select(visual => visual.Location!.SourceBlockIndex));
        if (result != null) {
            Assert.Equal(new int?[] { 1, 2 }, result.Pages.Select(page => page.Location.SourceBlockIndex));
            Assert.Equal(new int?[] { 1, 2 }, result.Metadata.Where(entry => entry.Category == "pdf.annotation" && entry.Value == "Readable annotation")
                .Select(entry => entry.Location!.SourceBlockIndex));
            if (!lean) Assert.Equal(new int?[] { 1, 2 }, result.Visuals.Select(visual => visual.Location!.SourceBlockIndex));
            Assert.Contains(result.Metadata, entry => entry.Category == "pdf.page-outcome" && entry.Attributes["selectionIndex"] == "0");
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LoadedInspectionKeepsNativeSelectionsWithoutReopeningItsSource(bool lean) {
        PdfInspectionReadResult inspection = PdfDocument.Load(BuildMixedPdf(decoratedReadablePage: true)).ReadForInspection(
            new PdfLoadOptions { PageSelection = PdfPageSelection.From(2, 1, 1) });
        OfficeDocumentReadResult result = inspection.ReadDocument("not-an-input-file.pdf", new ReaderOptions {
            Projection = lean ? ReaderProjection.TextAndMetadata : ReaderProjection.Full
        });
        Assert.Equal(new int?[] { 1, 2 }, result.Pages.Select(page => page.Location.SourceBlockIndex));
        Assert.Contains(result.Metadata, entry => entry.Category == "pdf.source.structure" && entry.Name == "HeaderVersion" && entry.Value == "1.7");
        Assert.Contains(result.Metadata, entry => entry.Category == "pdf.source.structure" && entry.Name == "SourceFormFieldCount" && entry.Value == "0");
        Assert.Contains(result.Metadata, entry => entry.Category == "pdf.source.structure" && entry.Name == "SourceFormWidgetCount" && entry.Value == "0");
        Assert.Equal(new int?[] { 1, 2 }, result.Chunks.Where(chunk => chunk.Warnings == null || chunk.Warnings.Count == 0)
            .Select(chunk => chunk.Location.SourceBlockIndex));
        Assert.Contains(result.Diagnostics, item => item.Code == "pdf-page-unmappable-text" && item.Location?.SourceBlockIndex == 0
            && item.IsRecoverable == true && item.Attributes["disposition"] == "Skipped");
        Assert.Equal(new int?[] { 1, 2 }, result.Metadata.Where(entry => entry.Category == "pdf.annotation" && entry.Value == "Readable annotation")
            .Select(entry => entry.Location!.SourceBlockIndex));
        if (!lean) Assert.Equal(new int?[] { 1, 2 }, result.Visuals.Select(visual => visual.Location!.SourceBlockIndex));
        if (!lean) {
            Assert.Equal(new int?[] { 1, 2 }, result.Blocks.Select(block => block.Location.SourceBlockIndex).Distinct());
            Assert.All(result.Pages[0].Blocks, block => Assert.Equal(1, block.Location.SourceBlockIndex));
            Assert.All(result.Pages[1].Blocks, block => Assert.Equal(2, block.Location.SourceBlockIndex));
        }
        Assert.Throws<ArgumentException>(() => inspection.ReadDocument(pdfOptions: new ReaderPdfOptions {
            ReadOptions = new PdfLoadOptions { PageSelection = PdfPageSelection.From(1) }
        }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageOccurrencesKeepSelectionIdentityThroughTheVisualReader(bool chunkByPage) {
        var reader = new OfficeDocumentReaderBuilder().AddPdfHandler(new ReaderPdfOptions {
            ChunkByPage = chunkByPage,
            ReadOptions = new PdfLoadOptions { PageSelection = PdfPageSelection.From(2, 1, 1) }
        }).Build();
        ReaderVisual[] visuals = reader.ReadVisuals(BuildMixedPdf(decoratedReadablePage: true), "selected.pdf",
            new ReaderOptions { Inspection = new ReaderInspectionOptions() }).ToArray();
        Assert.Equal(new int?[] { 1, 2 }, visuals.Select(visual => visual.Location!.SourceBlockIndex));
        Assert.Contains("selection-0001", visuals[0].Location!.BlockAnchor);
        Assert.Contains("selection-0002", visuals[1].Location!.BlockAnchor);
    }

    [Theory]
    [InlineData("input", false)]
    [InlineData("input", true)]
    [InlineData("pages", false)]
    [InlineData("pages", true)]
    [InlineData("content", false)]
    [InlineData("content", true)]
    public async Task NativePdfQuotasStayTerminalAcrossDocumentAndAggregateRoutes(string quota, bool inspection) {
        byte[] bytes = BuildMixedPdf();
        var limits = quota == "input" ? new PdfReadLimits { MaxInputBytes = bytes.Length - 1 }
            : quota == "pages" ? new PdfReadLimits { MaxPages = 1 }
            : new PdfReadLimits { MaxContentOperations = 1 };
        var reader = new OfficeDocumentReaderBuilder().AddPdfHandler(new ReaderPdfOptions {
            LoadOptions = new PdfLoadOptions { Limits = limits },
            ReadOptions = new PdfLoadOptions { PageSelection = PdfPageSelection.From(1) }
        }).Build();
        var options = new ReaderOptions { Inspection = inspection ? new ReaderInspectionOptions() : null };
        ReaderAdmissionLimitException direct = Assert.Throws<ReaderAdmissionLimitException>(() => reader.ReadDocument(bytes, "bounded.pdf", options));
        PdfReadLimitException native = Assert.IsType<PdfReadLimitException>(direct.InnerException);
        Assert.Equal(native.Limit, direct.Maximum); Assert.Equal(native.Actual, direct.Actual);
        string folder = Path.Combine(Path.GetTempPath(), "reader-pdf-terminal-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try {
            string path = Path.Combine(folder, "bounded.pdf"); File.WriteAllBytes(path, bytes);
            Assert.Throws<ReaderAdmissionLimitException>(() => reader.ReadFolderDetailed(folder, options: options));
            ReaderAdmissionLimitException batch = await Assert.ThrowsAsync<ReaderAdmissionLimitException>(() => reader.ReadDocumentsDetailedAsync(new[] { path }, options));
            Assert.Equal(direct.LimitName, batch.LimitName);
            Assert.IsType<PdfReadLimitException>(batch.InnerException);
        } finally { Directory.Delete(folder, true); }
    }

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
        ReaderAdmissionLimitException parser = Assert.Throws<ReaderAdmissionLimitException>(() => reader.ReadDocument(bytes, "mixed.pdf", inspection));
        Assert.Equal(PdfReadLimitKind.Pages, Assert.IsType<PdfReadLimitException>(parser.InnerException).Kind);
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

    private static byte[] BuildMixedPdf(bool fallbackFont = false, bool decoratedReadablePage = false) {
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
        if (decoratedReadablePage) {
            objects[2] = "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 5 0 R >> /XObject << /Im1 10 0 R >> >> /Contents 7 0 R /Annots [11 0 R] >>";
            objects[6] = Stream("BT /F1 12 Tf 20 100 Td (Readable native text) Tj ET q 1 0 0 1 0 0 cm /Im1 Do Q");
            objects = objects.Concat(new[] {
                "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\nabc\nendstream",
                "<< /Type /Annot /Subtype /Link /Rect [20 30 80 50] /Contents (Readable annotation) /A << /S /URI /URI (https://example.test/readable-page) >> >>"
            }).ToArray();
        }
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
