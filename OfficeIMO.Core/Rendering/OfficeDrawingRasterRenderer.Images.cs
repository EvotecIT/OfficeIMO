using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    internal static void RenderImage(
        OfficeRasterCanvas canvas,
        OfficeDrawingImage drawingImage,
        double scale,
        IOfficeRasterImageCodec? imageCodec,
        long maximumRasterPixels,
        System.Threading.CancellationToken cancellationToken,
        string? diagnosticSource = null,
        ICollection<OfficeImageExportDiagnostic>? diagnosticSink = null) {
        (double targetWidth, double targetHeight) = GetImageTargetSize(canvas, drawingImage.Projection, scale);
        if (TryDecodeImage(
                drawingImage.EncodedBytes,
                drawingImage.ContentType,
                targetWidth,
                targetHeight,
                imageCodec,
                canvas.TextShapingProvider,
                canvas.TextShapingLanguage,
                diagnosticSink ?? canvas.DiagnosticSink,
                diagnosticSource ?? canvas.DiagnosticSource,
                canvas.TransformedTextBudget,
                maximumRasterPixels,
                cancellationToken,
                out OfficeRasterImage? image) &&
            image != null) {
            if (drawingImage.Opacity < 1D) {
                canvas.ChargeIntermediateSurfacePixels((long)image.Width * image.Height, maximumRasterPixels);
                image = ApplyImageOpacity(image, drawingImage.Opacity);
            }

            canvas.DrawImage(image, drawingImage.Projection.Scale(scale), drawingImage.Interpolate);
        }
    }

    private static (double Width, double Height) GetImageTargetSize(
        OfficeRasterCanvas canvas, OfficeImageProjection projection, double scale) {
        (double axisX, double axisY) = GetEffectAxisScales(
            projection.CreateFrameTransform().CreateDestinationTransform(),
            canvas.CoordinateScaleX, canvas.CoordinateScaleY);
        return (projection.Width * scale * axisX / projection.SourceWidth,
            projection.Height * scale * axisY / projection.SourceHeight);
    }

    private static bool TryDecodeImage(
        byte[] bytes,
        string? contentType,
        double targetWidth,
        double targetHeight,
        IOfficeRasterImageCodec? imageCodec,
        IOfficeTextShapingProvider? textShapingProvider,
        string? textShapingLanguage,
        ICollection<OfficeImageExportDiagnostic>? diagnosticSink,
        string? diagnosticSource,
        OfficeRasterTransformedTextBudget transformedTextBudget,
        long maximumRasterPixels,
        System.Threading.CancellationToken cancellationToken,
        out OfficeRasterImage? image) {
        // Inspect the selected output before managed decode allocates it. The
        // remaining shared budget may be smaller than the per-image limit.
        long reservedPixels = 0L;
        long remainingPixels = maximumRasterPixels - transformedTextBudget.IntermediatePixels;
        var decodeOptions = new OfficeRasterDecodeOptions {
            ImageCodec = imageCodec is OfficeRasterImageFallbackCodec fallbackCodec ? fallbackCodec.SourceCodec : imageCodec,
            MaximumDecodedPixels = Math.Min(maximumRasterPixels, OfficeRasterGuards.MaximumPixels),
            MaximumInspectionWorkPixels = Math.Max(1L, Math.Min(remainingPixels, OfficeRasterGuards.MaximumPixels)),
            CancellationToken = cancellationToken
        };
        bool identifiedManagedRaster = OfficeImageReader.TryIdentifyByContent(bytes, null, out OfficeImageInfo identified) &&
            (identified.Format == OfficeImageFormat.Png || identified.Format == OfficeImageFormat.Jpeg ||
             identified.Format == OfficeImageFormat.Bmp || identified.Format == OfficeImageFormat.Webp ||
             identified.Format == OfficeImageFormat.Gif || identified.Format == OfficeImageFormat.Tiff ||
             identified.Format == OfficeImageFormat.Avif);
        if (identifiedManagedRaster) {
            if (!OfficeRasterImageDecoder.IsWithinPixelLimit(identified.Width, identified.Height, maximumRasterPixels)) {
                image = null;
                AddImageDecodeOmission(diagnosticSink, diagnosticSource,
                    "The embedded image dimensions exceed the configured raster limit.");
                if (imageCodec is RequiredImageCodec) throw new NotSupportedException(
                    "Raster rendering cannot decode the image within the raster limit.");
                return false;
            }
            // Identification reports the first TIFF page and the output canvas
            // for other managed formats. Reserve before inspection, which may
            // validate GIF frames or decode WebP pixels.
            reservedPixels = (long)identified.Width * identified.Height;
            transformedTextBudget.ChargeIntermediateSurfacePixels(reservedPixels, maximumRasterPixels);
        }
        bool decoded = false;
        image = null;
        OfficeRasterDecodeInfo decodeInfo;
        try {
            decoded = OfficeRasterImageDecoder.TryDecode(bytes, decodeOptions, out image, out decodeInfo) && image != null;
        } finally {
            if (!decoded && reservedPixels > 0L) transformedTextBudget.ReleaseIntermediateSurfacePixels(reservedPixels);
        }
        if (decoded && image != null) {
            if (decodeInfo.UsedCallerCodec && imageCodec is OfficeRasterImageFallbackCodec successfulFallback)
                successfulFallback.AddCallerCodecDiagnostic(contentType);
            if (decodeInfo.AnimationDiscarded || decodeInfo.FramesOrPagesDiscarded) {
                diagnosticSink?.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    OfficeImageExportDiagnosticCodes.SourceImageStaticFrameSelected,
                    decodeInfo.Diagnostic ?? "The selected static image does not retain animation or other frames/pages.",
                    diagnosticSource, OfficeConversionLossKind.Omission));
            }
            if (reservedPixels == 0L) {
                transformedTextBudget.ChargeIntermediateSurfacePixels((long)image.Width * image.Height, maximumRasterPixels);
            } else if ((long)image.Width * image.Height != reservedPixels) {
                long actualPixels = (long)image.Width * image.Height;
                if (actualPixels > reservedPixels) {
                    transformedTextBudget.ChargeIntermediateSurfacePixels(actualPixels - reservedPixels, maximumRasterPixels);
                } else {
                    transformedTextBudget.ReleaseIntermediateSurfacePixels(reservedPixels - actualPixels);
                }
            }
            return true;
        }
        // A placeholder is a representation of decode failure, not decoded source
        // pixels. Generate it only after validating the managed source container.
        if (identifiedManagedRaster && decodeInfo.Container != null && imageCodec is OfficeRasterImageFallbackCodec fallback) {
            int width = Math.Min(32, decodeInfo.Container.CanvasWidth);
            int height = Math.Min(32, decodeInfo.Container.CanvasHeight);
            transformedTextBudget.ChargeIntermediateSurfacePixels((long)width * height, maximumRasterPixels);
            image = fallback.CreateFallbackImage(contentType, width, height, decodeInfo.Diagnostic);
            return true;
        }
        // Managed raster providers, including caller-decoded WebP, run inside
        // the shared inspected boundary. Only the unsupported JPEG frame subset
        // has a separate fallback because managed inspection cannot describe it.
        bool callerCodecInputWithinLimit = bytes.Length <= decodeOptions.MaximumEncodedBytes;
        bool callerDecodedJpeg = callerCodecInputWithinLimit && identifiedManagedRaster && identified.Format == OfficeImageFormat.Jpeg &&
            OfficeImageReader.HasCompleteJpegPayload(bytes, cancellationToken,
                requireManagedFrame: false, validateMetadata: true) &&
            !OfficeImageReader.HasCompleteJpegPayload(bytes, cancellationToken,
                requireManagedFrame: true, validateMetadata: true);
        if (identifiedManagedRaster && !callerDecodedJpeg ||
            OfficeImageReader.HasWebpSignature(bytes) || OfficeImageReader.HasAvifSignature(bytes, cancellationToken) || !callerCodecInputWithinLimit) {
            AddImageDecodeOmission(diagnosticSink, diagnosticSource,
                decodeInfo.Diagnostic ?? "The embedded image could not be decoded within the configured limits.");
            if (imageCodec is RequiredImageCodec) throw new NotSupportedException(
                "Raster rendering cannot decode the image within the managed raster limits.");
            return false;
        }
        if (IsSvg(bytes, contentType) &&
            OfficeSvgDrawingReader.TryRead(bytes, out OfficeDrawing? vector, out int unsupportedFeatureCount) &&
            vector != null &&
            unsupportedFeatureCount == 0) {
            cancellationToken.ThrowIfCancellationRequested();
            double scaleX = Math.Max(1D, targetWidth) / vector.Width;
            double scaleY = Math.Max(1D, targetHeight) / vector.Height;
            double scale = Math.Max(scaleX, scaleY);
            double vectorPixels = Math.Ceiling(vector.Width * scaleX) * Math.Ceiling(vector.Height * scaleY);
            if (vectorPixels > long.MaxValue) {
                throw new OfficeImageExportLimitException(scale, long.MaxValue, maximumRasterPixels,
                    OfficeRasterImageEncoder.GetMaximumDimension(OfficeImageExportFormat.Png));
            }
            transformedTextBudget.ChargeIntermediateSurfacePixels((long)vectorPixels, maximumRasterPixels);
            // Plan and charge the actual projected axes. A uniform nested-vector
            // cap can both allocate an unrelated large surface and blur fine codes.
            image = RenderCore(vector, new OfficeDrawingRasterRenderOptions {
                Scale = scale,
                Background = OfficeColor.Transparent,
                ImageCodec = imageCodec,
                TextShapingProvider = textShapingProvider,
                TextShapingLanguage = textShapingLanguage,
                DiagnosticSink = diagnosticSink,
                DiagnosticSource = diagnosticSource,
                TransformedTextBudget = transformedTextBudget,
                MaximumRasterPixels = maximumRasterPixels,
                CancellationToken = cancellationToken
            }, scaleX, scaleY);
            return true;
        }
        var jpegFallback = callerDecodedJpeg ? imageCodec as OfficeRasterImageFallbackCodec : null;
        var sourceCodec = jpegFallback != null ? jpegFallback.SourceCodec : imageCodec;
        bool callerSucceeded = false;
        try {
            callerSucceeded = sourceCodec != null &&
                sourceCodec.TryDecode((byte[])bytes.Clone(), contentType, out image) && image != null;
        } catch (Exception exception) when (jpegFallback != null &&
            (exception is ArgumentException || exception is FormatException || exception is InvalidOperationException ||
             exception is System.IO.IOException || exception is NotSupportedException || exception is OverflowException)) {
            image = null;
        }
        cancellationToken.ThrowIfCancellationRequested();
        int expectedWidth = identified.Width;
        int expectedHeight = identified.Height;
        if (callerDecodedJpeg && OfficeImageOrientationNormalizer.TryRead(bytes, cancellationToken, out var orientation) &&
            orientation is >= OfficeImageOrientation.Transpose and <= OfficeImageOrientation.Rotate90CounterClockwise) {
            expectedWidth = identified.Height;
            expectedHeight = identified.Width;
        }
        if (callerSucceeded && image != null &&
            !OfficeRasterImageDecoder.IsWithinPixelLimit(image.Width, image.Height, maximumRasterPixels)) {
            AddImageDecodeOmission(diagnosticSink, diagnosticSource,
                "The caller-decoded image dimensions exceed the configured raster limit.");
        }
        if (callerSucceeded && image != null && (!callerDecodedJpeg || (image.Width == expectedWidth && image.Height == expectedHeight)) &&
            OfficeRasterImageDecoder.IsWithinPixelLimit(image.Width, image.Height, maximumRasterPixels)) {
            transformedTextBudget.ChargeIntermediateSurfacePixels((long)image.Width * image.Height, maximumRasterPixels);
            jpegFallback?.AddCallerCodecDiagnostic(contentType);
            return true;
        }
        image = null;
        if (jpegFallback != null) {
            int width = Math.Min(32, identified.Width), height = Math.Min(32, identified.Height);
            transformedTextBudget.ChargeIntermediateSurfacePixels((long)width * height, maximumRasterPixels);
            image = jpegFallback.CreateFallbackImage(contentType, width, height, decodeInfo.Diagnostic);
            return true;
        }
        return false;
    }

    private static void AddImageDecodeOmission(
        ICollection<OfficeImageExportDiagnostic>? diagnostics, string? source, string message) =>
        diagnostics?.Add(new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Warning,
            OfficeImageExportDiagnosticCodes.SourceImageDecodeOmitted,
            message, source, OfficeConversionLossKind.Omission));

    private static bool IsSvg(byte[] bytes, string? contentType) =>
        OfficeImageInfo.FromMimeType(contentType) == OfficeImageFormat.Svg ||
        (OfficeImageReader.TryIdentifyByContent(bytes, null, out OfficeImageInfo info) &&
         info.Format == OfficeImageFormat.Svg);

    private static OfficeRasterImage ApplyImageOpacity(OfficeRasterImage image, double opacity) {
        var result = new OfficeRasterImage(image.Width, image.Height);
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                byte alpha = (byte)Math.Round(pixel.A * opacity);
                result.SetPixel(x, y, OfficeColor.FromRgba(pixel.R, pixel.G, pixel.B, alpha));
            }
        }

        return result;
    }

}
