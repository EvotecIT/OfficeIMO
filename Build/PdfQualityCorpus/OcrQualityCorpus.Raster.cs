using OfficeIMO.Drawing;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.PdfQualityCorpus;

internal static partial class OcrQualityCorpus {
    private static async Task<OcrRequest> PrepareRasterAsync(byte[] bytes, OcrQualityLabel label, string output, int repetition, CancellationToken token) {
        // The pinned single-page TIFF fixtures use encodings beyond the managed decoder. Inspect metadata
        // and retain their original samples; native recognition must report first-page geometry only.
        if (label.File.EndsWith(".tif", StringComparison.OrdinalIgnoreCase) && label.Region == null && !label.Cleanup) {
            token.ThrowIfCancellationRequested();
            if (!OfficeImageReader.TryIdentifyByContent(bytes, label.File, out OfficeImageInfo info) ||
                info.Format != OfficeImageFormat.Tiff || info.Width <= 0 || info.Height <= 0 || (long)info.Width * info.Height > 20_000_000)
                throw new NotSupportedException("Qualification requires one bounded TIFF page.");
            return new OcrRequest { Payload = bytes, MediaType = "image/tiff", FileName = label.Id + ".tif",
                CandidateId = label.Id, PageNumber = 1, Language = label.Language,
                PixelWidth = info.Width, PixelHeight = info.Height };
        }
        OfficeRasterImage? image;
        if (label.File.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) {
            var region = label.Region;
            PdfScanPreview preview = await PdfDocument.Load(bytes).PreviewScanAsync(label.Page, new PdfOcrMergeOptions {
                Dpi = 300, ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(new PdfPageRange(label.Page, label.Page)) },
                Regions = region == null ? Array.Empty<PdfOcrPageRegion>() : new[] {
                    new PdfOcrPageRegion(label.Page, region[0], region[1], region[2], region[3]) },
                ScanProcessing = label.Cleanup ? new OfficeScanProcessingOptions() : null
            }, cancellationToken: token);
            bytes = preview.GetPreparedPng();
        }
        if (!OfficeRasterImageDecoder.TryDecode(bytes, new OfficeRasterDecodeOptions {
            MaximumEncodedBytes = 25 * 1024 * 1024, MaximumDecodedPixels = 20_000_000,
            CancellationToken = token
        }, out image, out _) || image == null) throw new NotSupportedException("The qualification raster did not decode.");
        if (!label.File.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) {
            if (label.Region != null) {
                int x = (int)Math.Floor(label.Region[0] * image.Width), y = (int)Math.Floor(label.Region[1] * image.Height);
                int right = Math.Min(image.Width, (int)Math.Ceiling((label.Region[0] + label.Region[2]) * image.Width));
                int bottom = Math.Min(image.Height, (int)Math.Ceiling((label.Region[1] + label.Region[3]) * image.Height));
                image = OfficeRasterResampler.Transform(image, OfficeTransform.Translate(-x, -y), right - x, bottom - y, cancellationToken: token);
            }
            if (label.Cleanup) image = OfficeScanProcessor.Process(image, new OfficeScanProcessingOptions(), token).Image;
        }
        bytes = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Png, options: null,
            maximumEncodedBytes: 25 * 1024 * 1024, cancellationToken: token);
        if (repetition == 1) await File.WriteAllBytesAsync(Path.Combine(output, label.Id + "-raster.png"), bytes, token);
        return new OcrRequest { Payload = bytes, MediaType = "image/png", FileName = label.Id + ".png",
            CandidateId = label.Id, PageNumber = label.Page, Language = label.Language,
            PixelWidth = image.Width, PixelHeight = image.Height };
    }

    private static bool GeometryWithinRaster(OcrResult result, OcrRequest request) {
        OcrTextSpan[] words = result.Spans.Where(span => span.Level == OcrTextSpanLevel.Word).ToArray();
        return words.Length > 0 && words.All(span => (!span.PageNumber.HasValue || span.PageNumber == 1) &&
            span.CoordinateUnit == OcrCoordinateUnit.Pixels && span.Region != null &&
            double.IsFinite(span.Region.X) && double.IsFinite(span.Region.Y) && double.IsFinite(span.Region.Width) && double.IsFinite(span.Region.Height) &&
            span.Region.X >= 0 && span.Region.Y >= 0 && span.Region.Width > 0 && span.Region.Height > 0 &&
            span.Region.X + span.Region.Width <= request.PixelWidth && span.Region.Y + span.Region.Height <= request.PixelHeight);
    }
}
