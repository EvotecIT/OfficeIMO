using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Ocr;

namespace OfficeIMO.Pdf.Ocr;

internal static partial class PdfOcr {
    private sealed class PreparedPage {
        internal PreparedPage(double width, double height) { SourceWidth = width; SourceHeight = height; }
        internal double SourceWidth { get; }
        internal double SourceHeight { get; }
        internal OfficeTransform PointsToSource { get; set; } = OfficeTransform.Identity;
        internal OfficeScanProcessingReport? Report { get; set; }
        internal List<string> Diagnostics { get; } = new List<string>();
    }

    private static async Task<PreparedPage> PreparePageAsync(OcrRequest request, OcrEngineExecution engine,
        PdfOcrMergeOptions options, CancellationToken token) {
        var prepared = new PreparedPage(request.Region!.Width, request.Region.Height);
        if (options.ScanProcessing == null && !options.DetectOrientation) return prepared;
        int detectedTurns = 0;
        if (options.DetectOrientation) {
            if (!engine.Capabilities.SupportsOrientationDetection) {
                prepared.Diagnostics.Add("ocr-orientation-unsupported: The provider does not detect orientation; retained the source orientation.");
            } else {
                var orientationRequest = new OcrRequest {
                    Operation = OcrOperation.DetectOrientation, Payload = request.Payload, MediaType = request.MediaType,
                    FileName = request.FileName, SourceId = request.SourceId, SourceName = request.SourceName,
                    CandidateId = request.CandidateId, CandidateKind = request.CandidateKind, PageNumber = request.PageNumber,
                    PixelWidth = request.PixelWidth, PixelHeight = request.PixelHeight, Region = request.Region,
                    RegionCoordinateUnit = request.RegionCoordinateUnit, Language = request.Language, ProviderOptions = request.ProviderOptions
                };
                OcrResult detection = await engine.RecognizeAsync(orientationRequest, options.ProviderTimeout, token).ConfigureAwait(false);
                // Use the same diagnostic count, length, metadata, and provider-result bounds as recognition.
                ProjectedOcrResult projected = ProjectResult(detection, request, engine.Id, options, token);
                prepared.Diagnostics.AddRange(projected.Diagnostics);
                OcrOrientationResult? orientation = detection.Orientation;
                if (orientation != null && orientation.ClockwiseRotationDegrees is 0 or 90 or 180 or 270 &&
                    IsFinite(orientation.Confidence) && orientation.Confidence >= options.MinimumOrientationConfidence && orientation.Confidence <= 1D) {
                    detectedTurns = orientation.ClockwiseRotationDegrees / 90;
                    prepared.Diagnostics.Add("ocr-orientation: Accepted provider clockwise correction " + orientation.ClockwiseRotationDegrees + " degrees.");
                } else {
                    prepared.Diagnostics.Add("ocr-orientation-inconclusive: Retained source orientation because provider evidence was missing, invalid, or below the confidence threshold.");
                }
            }
        }
        if (options.ScanProcessing == null && detectedTurns == 0) return prepared;
        OfficeScanProcessingOptions scanOptions = options.ScanProcessing?.Clone() ?? new OfficeScanProcessingOptions {
            Deskew = false, NormalizeBackground = false, ColorMode = OfficeScanColorMode.PreserveColor
        };
        if (scanOptions.ClockwiseQuarterTurns < 0 || scanOptions.ClockwiseQuarterTurns > 3)
            throw new ArgumentOutOfRangeException(nameof(scanOptions.ClockwiseQuarterTurns));
        scanOptions.ClockwiseQuarterTurns = (scanOptions.ClockwiseQuarterTurns + detectedTurns) % 4;
        int originalWidth = request.PixelWidth!.Value, originalHeight = request.PixelHeight!.Value;
        try {
            // Reject optional cleanup before decoding when the caller's processing budget cannot hold its input.
            long sourcePixels = (long)originalWidth * originalHeight;
            if (sourcePixels > scanOptions.MaximumPixels || sourcePixels * 4L > scanOptions.MaximumWorkingBytes) {
                prepared.Diagnostics.Add("ocr-scan-limit: Retained the original rendered image because its decoded input exceeds the scan-processing budget.");
                return prepared;
            }
            var decodeOptions = new OfficeRasterDecodeOptions {
                MaximumDecodedPixels = scanOptions.MaximumPixels, CancellationToken = token
            };
            if (!OfficeRasterImageDecoder.TryDecode(request.Payload, decodeOptions, out OfficeRasterImage? original, out _) || original == null)
                throw new NotSupportedException("The rendered PNG could not be decoded for scan processing.");
            OfficeScanProcessingResult processed = OfficeScanProcessor.Process(original, scanOptions, token);
            byte[] payload = OfficeRasterImageEncoder.Encode(processed.Image, OfficeImageExportFormat.Png,
                options: null, maximumEncodedBytes: options.MaxRenderedBytesPerPage, cancellationToken: token);
            prepared.Report = processed.Report;
            prepared.PointsToSource = OfficeTransform.Scale(originalWidth / prepared.SourceWidth, originalHeight / prepared.SourceHeight)
                .Then(processed.Report.ProcessedToSource)
                .Then(OfficeTransform.Scale(prepared.SourceWidth / originalWidth, prepared.SourceHeight / originalHeight));
            request.Payload = payload;
            request.PixelWidth = processed.Image.Width; request.PixelHeight = processed.Image.Height;
            request.Region = new OcrRegion {
                Width = processed.Image.Width * prepared.SourceWidth / originalWidth,
                Height = processed.Image.Height * prepared.SourceHeight / originalHeight
            };
        } catch (OfficeScanProcessingLimitException exception) {
            prepared.Diagnostics.Add("ocr-scan-limit: Retained the original rendered image. " + exception.Message);
        }
        if (prepared.Diagnostics.Count > options.MaxDiagnosticsPerPage)
            throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, options.MaxDiagnosticsPerPage, prepared.Diagnostics.Count);
        EnsureCharacters(prepared.Diagnostics, options.MaxDiagnosticCharactersPerPage);
        return prepared;
    }

    private static PdfSelectionQuad MapWordGeometry(double x, double y, double width, double height, PreparedPage? prepared) {
        OfficeTransform transform = prepared?.PointsToSource ?? OfficeTransform.Identity;
        PdfSelectionPoint Map(double px, double py) {
            OfficePoint point = transform.TransformPoint(new OfficePoint(px, py));
            return new PdfSelectionPoint(point.X, point.Y);
        }
        return new PdfSelectionQuad(Map(x, y), Map(x + width, y), Map(x + width, y + height), Map(x, y + height));
    }
}
