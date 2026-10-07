using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Ocr;

namespace OfficeIMO.Pdf.Ocr;

internal static partial class PdfOcr {
    internal static async Task<PdfSearchableOcrReview> RecognizeImageAsync(
        PdfImageDocumentSource image, IOcrEngine engine, PdfOcrMergeOptions? options,
        CancellationToken cancellationToken) {
        Guard.NotNull(image, nameof(image));
        Guard.NotNull(engine, nameof(engine));
        cancellationToken.ThrowIfCancellationRequested();
        PdfOcrMergeOptions effective = options?.Clone() ?? new PdfOcrMergeOptions { ReconstructLayout = true };
        effective.Validate();
        int[] selected = effective.GetSelectedPages(1);
        if (selected.Length != 1 || selected[0] != 1)
            throw new ArgumentException("Standalone image recognition requires page 1 exactly once.", nameof(options));
        if (image.EncodedByteLength > effective.MaxRenderedBytesPerPage)
            throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, effective.MaxRenderedBytesPerPage, image.EncodedByteLength);

        byte[] bytes = image.GetBytes(cancellationToken);
        var decodeOptions = new OfficeRasterDecodeOptions {
            MaximumDecodedPixels = Math.Min(effective.MaxPixelsPerPage, 50_000_000L),
            FrameLossPolicy = OfficeRasterFrameLossPolicy.RejectMultipleFrames,
            ImageCodec = effective.ImageCodec,
            CancellationToken = cancellationToken
        };
        decodeOptions.MaximumEncodedBytes = (int)Math.Min(effective.MaxRenderedBytesPerPage, decodeOptions.MaximumEncodedBytes);
        if (!OfficeRasterImageDecoder.TryDecode(bytes, decodeOptions, out _, out OfficeRasterDecodeInfo decoded))
            throw new NotSupportedException(decoded.Diagnostic ?? "The image is malformed, unsupported, or exceeds recognition limits.");

        // Reuse the composition owner's orientation and embedding decisions. No page rendering or resampling
        // occurs before recognition; the PDF snapshot supplies geometry and subsequent searchable output only.
        PdfDocument.PreparedImage preparedImage = PdfDocument.PrepareImageDocumentSource(bytes, cancellationToken);
        OfficeImageInfo info = preparedImage.Info;
        PdfDocument source = PdfDocument.CreateFromImages(
            new[] { new PdfImageDocumentSource(preparedImage.Data, image.Name) },
            new PdfImageDocumentOptions { MaximumEncodedImageBytes = effective.MaxRenderedBytesPerPage }, cancellationToken);
        PdfReadDocument readDocument = PdfReadDocument.Open(source.GetBytesForOperation(cancellationToken), source.ReadOptions, cancellationToken);
        PdfDocumentReadResult native = PdfDocumentReadEngine.Read(readDocument, effective.ReadOptions,
            out IReadOnlyList<PdfUnderstandingPageResult> analyses, cancellationToken);
        PdfLogicalPage page = native.Pages.Single();
        (double width, double height) = page.GetVisualPageSize();
        OcrEngineExecution execution = OcrEngineRunner.CreateExecution(engine);
        byte[] payload = preparedImage.Data;
        string mediaType = info.MimeType;
        if (!execution.Capabilities.SupportsMediaType(mediaType)) {
            EnsurePngSupport(execution.Capabilities);
            if (!OfficeRasterImageDecoder.TryDecode(payload, decodeOptions, out OfficeRasterImage? raster, out _) || raster == null)
                throw new NotSupportedException("The image could not be normalized to the provider's supported PNG input.");
            payload = OfficeRasterImageEncoder.Encode(raster, OfficeImageExportFormat.Png, options: null,
                maximumEncodedBytes: effective.MaxRenderedBytesPerPage, cancellationToken: cancellationToken);
            mediaType = "image/png";
        }
        var request = new OcrRequest {
            Payload = payload, MediaType = mediaType,
            FileName = "image-page-1" + (mediaType == "image/jpeg" ? ".jpg" : ".png"),
            SourceName = effective.SourceName ?? image.Name, SourceId = effective.SourceId,
            CandidateId = "image-page-1", CandidateKind = "image", PageNumber = 1,
            PixelWidth = info.Width, PixelHeight = info.Height,
            Region = new OcrRegion { Width = width, Height = height }, RegionCoordinateUnit = OcrCoordinateUnit.Points,
            Language = effective.Language, ProviderOptions = effective.ProviderOptions
        };
        PreparedPage prepared = await PreparePageAsync(request, execution, effective, cancellationToken).ConfigureAwait(false);
        if (prepared.HasGeometryTransform) {
            prepared.RecognitionWidth = request.Region!.Width;
            prepared.RecognitionHeight = request.Region.Height;
        }
        OcrResult recognized = await execution.RecognizeAsync(request, effective.ProviderTimeout,
            new OcrResultCaptureLimits(effective.MaxOcrSpansPerPage, effective.MaxDiagnosticsPerPage, 0), cancellationToken).ConfigureAwait(false);
        ProjectedOcrResult projected = ProjectResult(recognized, request, execution.Id, effective, cancellationToken, prepared);
        if (prepared.Diagnostics.Count > 0) {
            string[] diagnostics = prepared.Diagnostics.Concat(projected.Diagnostics).ToArray();
            if (diagnostics.Length > effective.MaxDiagnosticsPerPage)
                throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, effective.MaxDiagnosticsPerPage, diagnostics.Length);
            EnsureCharacters(diagnostics, effective.MaxDiagnosticCharactersPerPage);
            projected = new ProjectedOcrResult(projected.Words, Array.AsReadOnly(diagnostics), projected.Provider,
                projected.Model, projected.Language,
                Array.AsReadOnly(prepared.ProviderDiagnostics.Concat(projected.ProviderDiagnostics).ToArray()));
        }
        PdfOcrPageMergeResult merged = MergePage(page, Array.Empty<PdfSelectionQuad>(), projected, effective, cancellationToken, prepared);
        PdfOcrMergeResult result = BuildMergedResult(readDocument, native, analyses, new[] { merged }, effective, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return new PdfSearchableOcrReview(source, effective, result);
    }
}
