using OfficeIMO.Ocr;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf.Ocr;

public static partial class PdfOcrExtensions {
    /// <summary>Detects whole-page corrective quarter-turns on a private snapshot without changing the PDF.</summary>
    /// <remarks>Uses the shared OCR provider timeout, serialization and diagnostic budgets. Regions and scan
    /// preprocessing are rejected because their orientation does not establish the whole page's orientation.</remarks>
    public static Task<PdfPageOrientationReview> PrepareOrientationReviewAsync(this PdfDocument document,
        IOcrEngine engine, PdfOcrMergeOptions? options = null, CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document)); Guard.NotNull(engine, nameof(engine));
        var effective = options?.Clone() ?? new PdfOcrMergeOptions(); effective.Validate();
        effective.AwaitProviderSettlement = true;
        if (effective.Regions.Count != 0 || effective.ScanProcessing != null || effective.Perspective != null)
            throw new ArgumentException("Page orientation review requires whole pages without scan preprocessing.", nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        var source = PdfDocument.Load(document.GetBytesForOperation(cancellationToken), document.ReadOptions);
        NormalizeReviewPageSelection(effective, source.Inspect(source.ReadOptions, cancellationToken).PageCount);
        return PdfOcr.PrepareOrientationReviewAsync(source, engine, effective, cancellationToken);
    }
}

internal static partial class PdfOcr {
    internal static async Task<PdfPageOrientationReview> PrepareOrientationReviewAsync(PdfDocument source,
        IOcrEngine engine, PdfOcrMergeOptions options, CancellationToken token) {
        var execution = OcrEngineRunner.CreateExecution(engine);
        EnsurePngSupport(execution.Capabilities);
        var read = PdfReadDocument.Open(source.GetBytesForOperation(token), source.ReadOptions, token);
        int[] selected = options.GetSelectedPages(read.Pages.Count);
        var pages = new List<PdfPageOrientationProposal>();
        foreach (int page in selected) {
            token.ThrowIfCancellationRequested();
            int original = read.Pages[page - 1].GetRotationDegrees();
            if (!execution.Capabilities.SupportsOrientationDetection) {
                pages.Add(new(page, original, null, null, false, execution.Id,
                    new[] { "ocr-orientation-unsupported: The provider does not detect orientation." }));
                continue;
            }
            var render = RenderOcrPage(read, page, new PdfPageRenderOptions {
                Format = PdfPageRenderFormat.Png, Dpi = options.Dpi, ImageCodec = options.ImageCodec,
                MaxPixelsPerPage = options.MaxPixelsPerPage, MaxOutputBytesPerPage = options.MaxRenderedBytesPerPage,
                MaxDiagnosticsPerPage = options.MaxDiagnosticsPerPage,
                MaxDiagnosticCharactersPerPage = options.MaxDiagnosticCharactersPerPage, ContinueOnError = false
            }, token);
            var (width, height) = read.Pages[page - 1].GetInteractionPageSize();
            var request = new OcrRequest {
                Operation = OcrOperation.DetectOrientation, Payload = render.Bytes!, MediaType = "image/png",
                PageNumber = page, PixelWidth = render.Width, PixelHeight = render.Height,
                Region = new OcrRegion { Width = width, Height = height }, RegionCoordinateUnit = OcrCoordinateUnit.Points,
                CandidateKind = "page", CandidateId = "pdf-page-" + page, FileName = "pdf-page-" + page + ".png",
                Language = options.Language, ProviderOptions = options.ProviderOptions,
                SourceName = options.SourceName, SourceId = options.SourceId
            };
            OcrResult result = await RecognizeWithLifetimeAsync(execution, request, options, token).ConfigureAwait(false);
            var projected = ProjectResult(result, request, execution.Id, options, token);
            var diagnostics = render.Diagnostics.Concat(projected.Diagnostics).ToArray();
            if (diagnostics.Length > options.MaxDiagnosticsPerPage)
                throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, options.MaxDiagnosticsPerPage, diagnostics.Length);
            EnsureCharacters(diagnostics, options.MaxDiagnosticCharactersPerPage);
            var evidence = result.Orientation;
            bool valid = evidence != null && evidence.ClockwiseRotationDegrees is 0 or 90 or 180 or 270 &&
                IsFinite(evidence.Confidence) && evidence.Confidence >= 0 && evidence.Confidence <= 1;
            pages.Add(new(page, original, valid ? evidence!.ClockwiseRotationDegrees : null,
                valid ? evidence!.Confidence : null,
                valid && evidence!.ClockwiseRotationDegrees != 0 && evidence.Confidence >= options.MinimumOrientationConfidence,
                projected.Provider ?? execution.Id, diagnostics));
        }
        token.ThrowIfCancellationRequested();
        return new PdfPageOrientationReview(source, options, pages);
    }
}
