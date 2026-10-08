using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfImageOcrTests {
    [Theory]
    [InlineData("english-ledger", false)]
    [InlineData("english-ledger", true)]
    [InlineData("english-columns", false)]
    public async Task IndependentImageReconstructsEditableContentAndSearchableReadback(string id, bool omitDensityMetadata) {
        string root = Path.Combine(AppContext.BaseDirectory, "EnglishLayout");
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "manifest.json")));
        JsonElement fixture = manifest.RootElement.GetProperty("cases").EnumerateArray().Single(item => item.GetProperty("id").GetString() == id);
        byte[] png = ReadVerified("scan.png");
        if (omitDensityMetadata) png = WithoutDensity(png);
        string tsv = Encoding.UTF8.GetString(ReadVerified("provider.tsv"));
        string[] expectedOrder = fixture.GetProperty("readingOrder").EnumerateArray().Select(item => item.GetString()!).ToArray();
        string[][] expectedTable = fixture.GetProperty("table").EnumerateArray()
            .Select(row => row.EnumerateArray().Select(cell => cell.GetString()!).ToArray()).ToArray();
        OfficeImageInfo dimensions = OfficeImageReader.Identify(png);
        var engine = new DelegateOcrEngine("recorded-independent-tesseract", (request, _) => {
            // The provider boundary must receive the original pixels even when the PDF render DPI differs.
            Assert.Equal(png, request.Payload);
            Assert.Equal(dimensions.Width, request.PixelWidth);
            Assert.Equal(dimensions.Height, request.PixelHeight);
            return Task.FromResult(TesseractTsvParser.Parse(tsv, "eng"));
        }, new OcrEngineCapabilities { SupportsWordSpans = true, SupportedMediaTypes = new[] { "image/png" } });
        var image = new PdfImageDocumentSource(png, id + ".png");
        PdfSearchableOcrReview review = await image.PrepareSearchableOcrAsync(engine,
            new PdfOcrMergeOptions { Dpi = 72, MinimumConfidence = 0, ReconstructLayout = true });
        Assert.Equal(Normalize(string.Join(" ", expectedOrder)), Normalize(review.Ocr.Text));
        Assert.True(review.Ocr.HasAcceptedOcrContent);
        using (JsonDocument exported = JsonDocument.Parse(review.Ocr.Document.ExportStructured(PdfStructuredExportFormat.Json))) {
            JsonElement page = exported.RootElement.GetProperty("pages")[0];
            Assert.Equal(Normalize(review.Ocr.Text), Normalize(page.GetProperty("text").GetString()!));
            if (expectedTable.Length == 0) Assert.Empty(page.GetProperty("tables").EnumerateArray());
            else Assert.Single(page.GetProperty("tables").EnumerateArray());
            if (expectedTable.Length > 0)
                Assert.Equal(expectedTable, page.GetProperty("tables")[0].GetProperty("rows").EnumerateArray()
                    .Select(row => row.EnumerateArray().Select(cell => cell.GetString()!).ToArray()).ToArray());
        }
        string markdown = review.Ocr.Document.ToMarkdown(new PdfLogicalMarkdownOptions { IncludeImagePlaceholders = false });
        if (expectedTable.Length > 0) {
            Assert.Single(Regex.Matches(markdown, "Return credit").Cast<Match>());
            Assert.Contains("| Return credit | 3 | ", markdown);
            Assert.True(markdown.IndexOf("Reviewed Service Ledger", StringComparison.Ordinal) < markdown.IndexOf("| Return credit", StringComparison.Ordinal));
            Assert.True(markdown.IndexOf("| Return credit", StringComparison.Ordinal) < markdown.IndexOf("Report complete.", StringComparison.Ordinal));
        }
        if (expectedTable.Length == 0) Assert.Empty(review.Ocr.Document.Tables);
        else {
            PdfLogicalTable table = Assert.Single(review.Ocr.Document.Tables);
            Assert.True(expectedTable.Length == table.Rows.Count,
                string.Join(Environment.NewLine, table.Rows.Select(row => string.Join(" | ", row))));
            for (int row = 0; row < expectedTable.Length; row++) Assert.Equal(expectedTable[row], table.Rows[row]);
            var preserved = review.Ocr.Document.ImportTablesToExcelDocumentResult();
            using (preserved.Value) {
                Assert.False(Assert.Single(preserved.Report.Entries).FirstRowUsedAsHeader);
                Assert.Equal(expectedTable.Length, preserved.Report.Entries[0].RowCount);
                Assert.Equal("Description", preserved.Value.Sheets[0].CellAt(2, 1).GetValue<string>());
            }
            var limited = review.Ocr.Document.ImportTablesToExcelDocumentResult(new PdfTablesToExcelOptions { UseFirstRowAsHeader = true, MaxRows = 1 });
            using (limited.Value) {
                var entry = Assert.Single(limited.Report.Entries);
                Assert.Equal(1, entry.RowCount);
                Assert.Equal(expectedTable.Length - 1, entry.TotalRowCount);
                Assert.True(entry.Truncated);
            }
            var imported = review.Ocr.Document.ImportTablesToExcelDocumentResult(new PdfTablesToExcelOptions { UseFirstRowAsHeader = true });
            Assert.True(Assert.Single(imported.Report.Entries).FirstRowUsedAsHeader);
            using var workbook = imported.Value;
            using var bytes = new MemoryStream();
            workbook.Save(bytes);
            bytes.Position = 0;
            using var reopened = ExcelDocument.Load(bytes);
            var sheet = Assert.Single(reopened.Sheets);
            Assert.Equal("Return credit", sheet.CellAt(3, 1).GetValue<string>());
            Assert.Equal(3D, sheet.CellAt(3, 2).GetValue<double>());
            Assert.Equal(-4.5D, sheet.CellAt(3, 3).GetValue<double>());
            Assert.Equal(-13.5D, sheet.CellAt(3, 4).GetValue<double>());
            Assert.Equal(119.88D, sheet.CellAt(2, 4).GetValue<double>());
        }
        using (var word = review.Ocr.Document.ToWordDocument()) {
            using var bytes = new MemoryStream();
            word.Save(bytes);
            bytes.Position = 0;
            using var reopened = WordDocument.Load(bytes);
            Assert.Equal(expectedTable.Length == 0 ? 0 : 1, reopened.Tables.Count);
            if (expectedTable.Length > 0) {
                Assert.Equal(expectedTable.Length, reopened.Tables[0].Rows.Count);
                Assert.Equal("Return credit", reopened.Tables[0].Rows[2].Cells[0].Paragraphs[0].Text);
            }
        }
        PdfSearchableOcrResult searchable = review.ApplyAll();
        Assert.Equal(Normalize(review.Ocr.Text), Normalize(PdfDocument.Load(searchable.Document.ToBytes()).Read().Text));
        Assert.Equal(png, image.GetBytes());

        byte[] ReadVerified(string suffix) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(root, id + "-" + suffix));
            using var hash = SHA256.Create();
            string actual = string.Concat(hash.ComputeHash(bytes).Select(value => value.ToString("x2", CultureInfo.InvariantCulture)));
            Assert.Equal(fixture.GetProperty("files").GetProperty(suffix).GetString(), actual);
            return bytes;
        }
    }

    [Fact]
    public async Task LargerCallerByteBudgetRetainsDecoderCeilingWithoutRejectingSmallImage() {
        int calls = 0;
        var engine = new DelegateOcrEngine("large-budget", (_, _) => { calls++; return Task.FromResult(new OcrResult()); });
        await new PdfImageDocumentSource(PdfPngTestImages.CreateRgbPng(10, 10)).ReadWithOcrAsync(engine,
            new PdfOcrMergeOptions { MaxRenderedBytesPerPage = 256L * 1024 * 1024 });
        Assert.Equal(1, calls);
    }

    [Theory]
    [InlineData("none")]
    [InlineData("crop")]
    [InlineData("cleanup")]
    [InlineData("orientation")]
    [InlineData("perspective")]
    public async Task JpegPreparationKeepsPayloadMetadataConsistent(string preparation) {
        byte[] jpeg = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(200, 100, OfficeColor.White), OfficeImageExportFormat.Jpeg);
        int calls = 0;
        var engine = new DelegateOcrEngine("image-preparation", (request, _) => {
            calls++;
            bool original = preparation == "none" || request.Operation == OcrOperation.DetectOrientation;
            Assert.Equal(original ? "image/jpeg" : "image/png", request.MediaType);
            Assert.EndsWith(original ? ".jpg" : ".png", request.FileName);
            Assert.Equal(request.MediaType, OfficeImageReader.Identify(request.Payload).MimeType);
            if (original) Assert.Equal(jpeg, request.Payload);
            return Task.FromResult(request.Operation == OcrOperation.DetectOrientation
                ? new OcrResult { Orientation = new OcrOrientationResult { ClockwiseRotationDegrees = 90, Confidence = 1 } }
                : new OcrResult());
        }, new OcrEngineCapabilities { SupportsOrientationDetection = true, SupportedMediaTypes = new[] { "image/png", "image/jpeg" } });
        var options = PreparationOptions(preparation);
        await new PdfImageDocumentSource(jpeg).ReadWithOcrAsync(engine, options);
        Assert.Equal(preparation == "orientation" ? 2 : 1, calls);
    }

    [Theory]
    [InlineData("crop")]
    [InlineData("cleanup")]
    public async Task PreparedPngIsRejectedBeforeUnsupportedProviderExecution(string preparation) {
        byte[] jpeg = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(200, 100, OfficeColor.White), OfficeImageExportFormat.Jpeg);
        int calls = 0;
        var engine = new DelegateOcrEngine("jpeg-only", (_, _) => { calls++; return Task.FromResult(new OcrResult()); },
            new OcrEngineCapabilities { SupportedMediaTypes = new[] { "image/jpeg" } });
        await Assert.ThrowsAsync<NotSupportedException>(() => new PdfImageDocumentSource(jpeg).ReadWithOcrAsync(engine, PreparationOptions(preparation)));
        Assert.Equal(0, calls);
    }

    private static PdfOcrMergeOptions PreparationOptions(string preparation) {
        var options = new PdfOcrMergeOptions();
        if (preparation == "crop") options.Regions = new[] { new PdfOcrPageRegion(1, 0, 0, .5, 1) };
        if (preparation == "cleanup") options.ScanProcessing = new() { Deskew = false, NormalizeBackground = false, ColorMode = OfficeScanColorMode.PreserveColor };
        if (preparation == "orientation") options.DetectOrientation = true;
        if (preparation == "perspective") options.Perspective = new() {
            TopLeft = new OfficePoint(0, 0), TopRight = new OfficePoint(1, 0),
            BottomRight = new OfficePoint(1, 1), BottomLeft = new OfficePoint(0, 1)
        };
        return options;
    }

    [Fact]
    public async Task InputLimitsAndCancellationStopBeforeProviderExecution() {
        byte[] png = PdfPngTestImages.CreateRgbPng(10, 10);
        int calls = 0;
        var engine = new DelegateOcrEngine("not-called", (_, _) => { calls++; return Task.FromResult(new OcrResult()); });
        var source = new PdfImageDocumentSource(png);
        await Assert.ThrowsAsync<PdfReadLimitException>(() => source.ReadWithOcrAsync(engine,
            new PdfOcrMergeOptions { MaxRenderedBytesPerPage = png.Length - 1 }));
        await Assert.ThrowsAsync<NotSupportedException>(() => source.ReadWithOcrAsync(engine,
            new PdfOcrMergeOptions { MaxPixelsPerPage = 99 }));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => source.ReadWithOcrAsync(engine, cancellationToken: cancellation.Token));
        Assert.Equal(0, calls);
    }

    [Fact]
    public async Task LowConfidenceAndRecoverableWarningsRemainVisibleWithoutInventedText() {
        var source = new PdfImageDocumentSource(PdfPngTestImages.CreateRgbPng(100, 100));
        var engine = new DelegateOcrEngine("uncertain", (_, _) => Task.FromResult(new OcrResult {
            Text = "incorrect", Spans = new[] { new OcrTextSpan {
                Text = "incorrect", Level = OcrTextSpanLevel.Word, Confidence = 0.1,
                CoordinateUnit = OcrCoordinateUnit.Pixels, Region = new OcrRegion { X = 10, Y = 10, Width = 50, Height = 10 }
            } }, Diagnostics = new[] { new OcrDiagnostic {
                Code = "provider-warning", Severity = OcrDiagnosticSeverity.Warning, IsRecoverable = true, Message = "Recognition is uncertain."
            } }
        }));
        var result = await source.ReadWithOcrAsync(engine);
        Assert.False(result.HasAcceptedOcrContent);
        Assert.Throws<InvalidOperationException>(() => result.RequireAcceptedOcrContent());
        Assert.Empty(result.Document.Tables);
        Assert.Equal(1, result.Pages[0].RejectedLowConfidenceCount);
        Assert.Contains(result.Pages[0].ProviderDiagnostics, diagnostic => diagnostic.Severity == OcrDiagnosticSeverity.Warning && diagnostic.IsRecoverable);
    }

    [Fact]
    public async Task NonRecoverableProviderFailureStopsRecognition() {
        var source = new PdfImageDocumentSource(PdfPngTestImages.CreateRgbPng(100, 100));
        var engine = new DelegateOcrEngine("failed", (_, _) => Task.FromResult(new OcrResult {
            Diagnostics = new[] { new OcrDiagnostic {
                Code = "provider-failure", Severity = OcrDiagnosticSeverity.Error, IsRecoverable = false, Message = "Private provider failure."
            } }
        }));
        var failure = await Assert.ThrowsAsync<OcrEngineExecutionException>(() => source.ReadWithOcrAsync(engine));
        Assert.Equal(OcrEngineFailureKind.NonRecoverableDiagnostic, failure.Kind);
    }

    [Fact]
    public async Task MultiFrameImageIsRejectedWithoutSilentlyDroppingFrames() {
        byte[] single = Convert.FromBase64String("R0lGODlhAQABAJAAAAAAAP///ywAAAAAAQABAAACAkwBADs=");
        // Retain the complete GIF header and palette, then append two valid image descriptors before its trailer.
        byte[] multiple = single.Take(single.Length - 1).Concat(single.Skip(19).Take(single.Length - 20)).Concat(new byte[] { 0x3b }).ToArray();
        Assert.True(OfficeRasterImageDecoder.TryDecode(multiple, new OfficeRasterDecodeOptions(), out _, out var decoded));
        Assert.Equal(2, decoded.FrameCount);
        int calls = 0;
        var engine = new DelegateOcrEngine("not-called", (_, _) => { calls++; return Task.FromResult(new OcrResult()); });
        await Assert.ThrowsAsync<NotSupportedException>(() => new PdfImageDocumentSource(multiple).ReadWithOcrAsync(engine));
        Assert.Equal(0, calls);
    }

    private static string Normalize(string value) => Regex.Replace(value.Normalize(), @"\s+", " ").Trim();

    private static byte[] WithoutDensity(byte[] png) {
        using var output = new MemoryStream();
        output.Write(png, 0, 8);
        for (int offset = 8; offset < png.Length;) {
            int length = (png[offset] << 24) | (png[offset + 1] << 16) | (png[offset + 2] << 8) | png[offset + 3];
            int chunkLength = length + 12;
            if (Encoding.ASCII.GetString(png, offset + 4, 4) != "pHYs") output.Write(png, offset, chunkLength);
            offset += chunkLength;
        }
        return output.ToArray();
    }
}
