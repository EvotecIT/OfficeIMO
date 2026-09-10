using OfficeIMO.Drawing;
using OfficeIMO.Ocr;
using System.Threading;

namespace OfficeIMO.Pdf.Ocr;

internal static partial class PdfOcr {
    private static void PrepareRegionAndPerspective(OcrRequest request, PdfOcrMergeOptions options, PreparedPage prepared, CancellationToken token) {
        PdfOcrPageRegion? region = options.Regions.FirstOrDefault(item => item.PageNumber == request.PageNumber);
        if (region == null && options.Perspective == null) return;
        long maximumPixels = Math.Min(options.MaxPixelsPerPage, options.Perspective?.MaximumPixels ?? options.ScanProcessing?.MaximumPixels ?? 20_000_000L);
        long maximumBytes = options.Perspective?.MaximumWorkingBytes ?? options.ScanProcessing?.MaximumWorkingBytes ?? 256L * 1024 * 1024;
        long sourcePixels = (long)request.PixelWidth!.Value * request.PixelHeight!.Value;
        if (sourcePixels > maximumPixels || sourcePixels * 8L > maximumBytes)
            throw new InvalidOperationException("The selected OCR preparation exceeds its raster budget. Reduce the render DPI.");
        if (!OfficeRasterImageDecoder.TryDecode(request.Payload, new OfficeRasterDecodeOptions {
            MaximumDecodedPixels = maximumPixels,
            CancellationToken = token
        }, out OfficeRasterImage? image, out _) || image == null)
            throw new NotSupportedException("The rendered page could not be decoded for region OCR.");
        double pointPerPixelX = request.Region!.Width / image.Width, pointPerPixelY = request.Region.Height / image.Height;
        if (region != null) {
            int left = (int)Math.Floor(region.X * image.Width), top = (int)Math.Floor(region.Y * image.Height);
            int right = Math.Min(image.Width, (int)Math.Ceiling((region.X + region.Width) * image.Width));
            int bottom = Math.Min(image.Height, (int)Math.Ceiling((region.Y + region.Height) * image.Height));
            image = OfficeRasterResampler.Transform(image, OfficeTransform.Translate(-left, -top), right - left, bottom - top, cancellationToken: token);
            double offsetX = left * pointPerPixelX, offsetY = top * pointPerPixelY;
            prepared.PreparedToOriginal = point => new OfficePoint(point.X + offsetX, point.Y + offsetY);
            prepared.Diagnostics.Add("ocr-region: Sent only the selected visual rectangle to the provider; geometry is mapped to the original page.");
        }
        if (options.Perspective != null) {
            // Include the retained original decoded page while the intermediate correction allocates its output.
            var perspective = options.Perspective.Clone();
            perspective.MaximumWorkingBytes = Math.Min(perspective.MaximumWorkingBytes, maximumBytes - sourcePixels * 4L);
            var corrected = OfficeScanProcessor.CorrectPerspective(image, perspective, token);
            OfficeScanPerspectiveMap mapping = corrected.Mapping;
            Func<OfficePoint, OfficePoint>? prior = prepared.PreparedToOriginal;
            prepared.PreparedToOriginal = point => {
                OfficePoint original = mapping.MapProcessedToSource(new OfficePoint(point.X / pointPerPixelX, point.Y / pointPerPixelY));
                original = new OfficePoint(original.X * pointPerPixelX, original.Y * pointPerPixelY);
                return prior == null ? original : prior(original);
            };
            image = corrected.Image;
            prepared.Diagnostics.Add("ocr-perspective: Corrected the explicitly selected page corners for recognition; original page samples remain unchanged.");
        }
        request.Payload = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Png, options: null, maximumEncodedBytes: options.MaxRenderedBytesPerPage, cancellationToken: token);
        request.PixelWidth = image.Width; request.PixelHeight = image.Height;
        request.Region = new OcrRegion { Width = image.Width * pointPerPixelX, Height = image.Height * pointPerPixelY };
    }
}