using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.DjVu;

namespace OfficeIMO.DjVu.Pdf;

internal static class DjVuPdfConversionEngine {
    internal static PdfDocumentConversionResult Convert(DjVuDocument document, DjVuToPdfOptions? options, CancellationToken token) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        token.ThrowIfCancellationRequested();
        var settings = (options ?? new DjVuToPdfOptions()).Clone();
        if (settings.OcrEngine != null) throw new InvalidOperationException("An explicit OCR engine requires ToPdfDocumentResultAsync or another asynchronous conversion entrypoint.");
        return Build(document, settings, Select(document, settings, token), null, token);
    }

    internal static async Task<PdfDocumentConversionResult> ConvertAsync(DjVuDocument document, DjVuToPdfOptions? options, CancellationToken token) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        token.ThrowIfCancellationRequested();
        var settings = (options ?? new DjVuToPdfOptions()).Clone();
        var pages = Select(document, settings, token);
        OfficeDocumentOcrExecutionResult? ocr = null;
        if (settings.OcrEngine != null) {
            int[] missing = pages.Where(p => p.GetText(token).Status == DjVuTextStatus.Absent || p.GetText(token).Status == DjVuTextStatus.Empty)
                .Select(p => p.Number).ToArray();
            var readerOptions = ReaderOptions.CreateSafeIngestion();
            readerOptions.ResourceLimits!.MaxAssets = settings.MaxPages;
            readerOptions.ResourceLimits.MaxAssetBytes = settings.MaxTotalImageBytes;
            var input = document.ToReadResult(new ReaderDjVuOptions {
                PageNumbers = missing, ImageMode = ReaderDjVuImageMode.MissingTextPages, RenderOptions = settings.RenderOptions,
                ReadOptions = document.ReadOptions, MaxPageImages = settings.MaxPages, MaxPageImageBytes = settings.MaxImageBytesPerPage,
                MaxTotalPageImageBytes = settings.MaxTotalImageBytes
            }, readerOptions, cancellationToken: token);
            ocr = await input.ApplyOcrAsync(settings.OcrEngine, settings.OcrOptions, token).ConfigureAwait(false);
        }
        return Build(document, settings, pages, ocr, token);
    }

    private static DjVuPage[] Select(DjVuDocument document, DjVuToPdfOptions settings, CancellationToken token) {
        var pages = DjVuPageSelection.Select(document, settings.PageNumbers, Math.Min(settings.MaxPages, settings.PdfOptions!.MaxGeneratedPages!.Value), token);
        if (pages.Length == 0) throw new ArgumentException("PDF conversion requires at least one selected page.", nameof(settings.PageNumbers));
        return pages;
    }

    private static PdfDocumentConversionResult Build(DjVuDocument document, DjVuToPdfOptions settings, DjVuPage[] pages,
        OfficeDocumentOcrExecutionResult? ocr, CancellationToken token) {
        var pdfReport = new PdfConversionReport();
        settings.PdfOptions!.UseContentStreamCompressionByDefault();
        settings.PdfOptions.ReportDiagnosticsTo(pdfReport, "OfficeIMO.DjVu.Pdf");
        var pdf = PdfDocument.Create(settings.PdfOptions);
        var pageReports = new List<DjVuPdfPageReport>();
        var diagnostics = new List<OfficeConversionFidelityDiagnostic>();
        var textBudget = new DjVuPdfTextBudget(settings);
        var navigation = DjVuPdfNavigation.Project(document, pages, diagnostics, token);
        var assets = (ocr?.Document.Assets ?? Array.Empty<OfficeDocumentAsset>()).Where(a => a.Location.Page.HasValue)
            .ToDictionary(a => a.Location.Page!.Value);
        var recognitions = (ocr?.Recognitions ?? Array.Empty<OfficeDocumentOcrRecognition>()).ToDictionary(r => r.CandidateId, StringComparer.Ordinal);
        long imageBytes = 0;
        foreach (var page in pages) {
            token.ThrowIfCancellationRequested();
            double width = page.DisplayWidth * 72.0 / page.Dpi, height = page.DisplayHeight * 72.0 / page.Dpi;
            byte[] payload;
            int imageWidth, imageHeight;
            if (assets.TryGetValue(page.Number, out var asset)) {
                payload = asset.PayloadBytes!; imageWidth = asset.Width!.Value; imageHeight = asset.Height!.Value;
                foreach (var diagnostic in ocr!.Document.Diagnostics.Where(d => d.Code.StartsWith("djvu.render.", StringComparison.Ordinal) && d.Location?.Page == page.Number))
                    Add(diagnostics, diagnostic.Code, diagnostic.Message, diagnostic.Code == "djvu.render.annotation-omitted"
                        ? OfficeConversionLossKind.Omission : OfficeConversionLossKind.Approximation, page.Number);
            } else {
                long remaining = settings.MaxTotalImageBytes - imageBytes;
                if (remaining <= 0) throw new DjVuResourceLimitException(nameof(DjVuToPdfOptions.MaxTotalImageBytes));
                var rendered = page.Render(settings.RenderOptions, token);
                payload = OfficeRasterImageEncoder.Encode(rendered.Image, OfficeImageExportFormat.Png, null,
                    Math.Min(settings.MaxImageBytesPerPage, remaining), token);
                imageWidth = rendered.Image.Width; imageHeight = rendered.Image.Height;
                diagnostics.AddRange(rendered.FidelityDiagnostics);
            }
            if (payload.LongLength > settings.MaxTotalImageBytes - imageBytes)
                throw new DjVuResourceLimitException(nameof(DjVuToPdfOptions.MaxTotalImageBytes));
            imageBytes += payload.LongLength;
            if (settings.RenderOptions.Dpi.HasValue && settings.RenderOptions.Dpi.Value < page.Dpi)
                Add(diagnostics, "djvu.pdf.image-downsampled", "The page image uses a lower resolution than the source.", OfficeConversionLossKind.Approximation, page.Number);
            var stored = page.GetText(token);
            if (stored.Status == DjVuTextStatus.Corrupt)
                Add(diagnostics, "djvu.pdf.stored-text-corrupt", stored.Diagnostic ?? "The stored text is corrupt and was not replaced with OCR.", OfficeConversionLossKind.Omission, page.Number);
            long before = textBudget.Characters;
            var textSource = DjVuPdfTextSource.None;
            OfficeDocumentRecognitionEvidence? evidence = null;
            pdf.Compose(builder => builder.Page(p => p.Size(width, height).Margin(0).Canvas(canvas => {
                canvas.Image(payload, 0, 0, width, height);
                foreach (var entry in navigation.Where(e => e.SourcePage == page.Number))
                    canvas.OutlineNavigation(entry.Title, entry.Level, 0, 0, entry.Uri, entry.Order);
                if (!settings.IncludeText) {
                    if (stored.Status == DjVuTextStatus.Present) Add(diagnostics, "djvu.pdf.stored-text-omitted", "Stored text was excluded by the conversion options.", OfficeConversionLossKind.Omission, page.Number);
                    return;
                }
                if (stored.Status == DjVuTextStatus.Present) {
                    DjVuPdfText.AddStored(canvas, page, stored, textBudget, diagnostics, token);
                    textSource = DjVuPdfTextSource.StoredText;
                } else if (recognitions.TryGetValue("djvu-page-" + page.Number + "-ocr", out var recognized)) {
                    if (DjVuPdfText.AddOcr(canvas, page, recognized.Result, imageWidth, imageHeight, textBudget, diagnostics, token)) {
                        textSource = DjVuPdfTextSource.Ocr;
                        evidence = ocr!.Document.Blocks.FirstOrDefault(b => b.Location.Page == page.Number && b.Recognition != null)?.Recognition;
                        Add(diagnostics, "djvu.pdf.ocr-text", "Searchable text was recognized by the explicitly supplied OCR engine.", OfficeConversionLossKind.Approximation, page.Number);
                    }
                }
            })));
            pageReports.Add(new DjVuPdfPageReport(page.Number, pageReports.Count + 1, stored.Status, textSource,
                imageWidth, imageHeight, width, height, textBudget.Characters - before, evidence));
        }
        if (ocr != null) foreach (var diagnostic in ocr.Diagnostics) Add(diagnostics, diagnostic.Code, diagnostic.Message,
            diagnostic.Severity == OfficeDocumentDiagnosticSeverity.Information ? OfficeConversionLossKind.None : OfficeConversionLossKind.Approximation,
            diagnostic.Location?.Page);
        token.ThrowIfCancellationRequested();
        return new PdfDocumentConversionResult(pdf, pdfReport).WithSourceConversionReport(
            new DjVuPdfConversionReport(document.SourceSha256, pageReports, diagnostics, ocr?.Report));
    }

    internal static void Add(List<OfficeConversionFidelityDiagnostic> diagnostics, string code, string message, OfficeConversionLossKind kind, int? page) =>
        diagnostics.Add(new OfficeConversionFidelityDiagnostic(code, message, kind, "OfficeIMO.DjVu.Pdf", page.HasValue ? "page " + page.Value : null));
}
