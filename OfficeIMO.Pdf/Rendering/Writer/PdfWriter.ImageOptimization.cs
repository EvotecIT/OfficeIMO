using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static void ApplyPlacementAwareImageOptimization(
        PageImage image,
        PdfOptions documentOptions,
        Dictionary<string, OfficeImageOptimizationResult> cache,
        System.Threading.CancellationToken cancellationToken) {
        PdfImageOptimizationOptions? options = documentOptions.ImageOptimizationSnapshot;
        if (options?.Enabled != true) return;
        if (!CanOptimizeImageFormat(image.Info.Format)) {
            ReportImageOptimization(documentOptions, image, "ImageOptimizationUnsupported", "Original image format was preserved.");
            return;
        }

        // W/H are the full source placement after the render plan has expanded for crop and fit.
        int targetWidth = ResolvePlacementPixelSize(image.W, options.TargetDpi);
        int targetHeight = ResolvePlacementPixelSize(image.H, options.TargetDpi);
        bool downsample = options.Mode != OfficeImageOptimizationMode.Recompress &&
            RequiresDownsampling(image.Info, targetWidth, targetHeight, options.DownsampleThreshold);
        bool recompress = options.Mode != OfficeImageOptimizationMode.Downsample && image.Info.Format == OfficeImageFormat.Jpeg;
        bool metadataPolicy = options.MetadataPolicy == OfficeImageMetadataPolicy.Strip ||
            (options.MetadataPolicy == OfficeImageMetadataPolicy.SelectiveCopy && options.MetadataSelection != OfficeImageMetadataKinds.All);
        if (!downsample && !recompress && !metadataPolicy) return;
        if (!downsample) {
            targetWidth = image.Info.Width;
            targetHeight = image.Info.Height;
        } else {
            double scale = Math.Min(1D, Math.Max(targetWidth / (double)image.Info.Width, targetHeight / (double)image.Info.Height));
            targetWidth = Math.Max(1, (int)Math.Ceiling(image.Info.Width * scale));
            targetHeight = Math.Max(1, (int)Math.Ceiling(image.Info.Height * scale));
        }

        string cacheKey = BuildPlacementOptimizationCacheKey(image.Data, targetWidth, targetHeight, options);
        if (!cache.TryGetValue(cacheKey, out OfficeImageOptimizationResult? result)) {
            try {
                result = OfficeImageOptimizer.Optimize(
                    image.Data,
                    new OfficeImageOptimizationRequest(targetWidth, targetHeight) {
                        Mode = options.Mode,
                        CancellationToken = cancellationToken,
                        ResamplingMode = options.ResamplingMode,
                        JpegQuality = options.JpegQuality,
                        MetadataPolicy = options.MetadataPolicy,
                        MetadataSelection = options.MetadataSelection,
                        KeepOriginalWhenNotSmaller = options.KeepOriginalWhenNotSmaller
                    });
            } catch (Exception exception) when (
                exception is ArgumentException ||
                exception is FormatException ||
                exception is InvalidOperationException ||
                exception is OverflowException) {
                ReportImageOptimization(documentOptions, image, "ImageOptimizationFailed", "Original image retained because optimization failed: " + exception.Message);
                return;
            }

            cache[cacheKey] = result;
        }

        if (!result.Changed) {
            ReportImageOptimization(documentOptions, image, "ImageOptimizationPreserved", "Original image retained: " + result.Status + ".");
            return;
        }
        bool removesMetadata = result.Metadata.HasLoss || result.Metadata.Stripped != OfficeImageMetadataKinds.None;
        if (removesMetadata && !options.AllowMetadataLoss) {
            ReportImageOptimization(documentOptions, image, "ImageOptimizationMetadataLoss", "Original image retained because the candidate would lose metadata.");
            return;
        }
        ReportImageOptimization(documentOptions, image, "ImageOptimized",
            "Encoded image changed from " + result.OriginalEncodedLength.ToString(CultureInfo.InvariantCulture) + " to " +
            result.FinalEncodedLength.ToString(CultureInfo.InvariantCulture) + " bytes; " +
            result.Final.Width.ToString(CultureInfo.InvariantCulture) + "x" + result.Final.Height.ToString(CultureInfo.InvariantCulture) + " pixels.");
        if (removesMetadata) {
            ReportImageOptimization(documentOptions, image, "ImageOptimizationMetadataRemoved",
                "Image metadata removed: lost " + result.Metadata.Lost + "; stripped " + result.Metadata.Stripped + ".",
                PdfConversionWarningSeverity.Warning);
        }
        image.Data = result.Bytes;
        image.Info = result.Final;
        image.PreparedStream = null;
    }

    private static void ReportImageOptimization(PdfOptions options, PageImage image, string code, string message,
        PdfConversionWarningSeverity severity = PdfConversionWarningSeverity.Information) =>
        options.AddLayoutDiagnostic(code, "image at " + image.X.ToString("R", CultureInfo.InvariantCulture) + "," +
            image.Y.ToString("R", CultureInfo.InvariantCulture), message, PdfLayoutDiagnosticKind.ImageOptimization,
            severity, image.X, image.Y, image.W, image.H);

    private static bool CanOptimizeImageFormat(OfficeImageFormat format) =>
        format == OfficeImageFormat.Png || format == OfficeImageFormat.Jpeg || format == OfficeImageFormat.Bmp ||
        format == OfficeImageFormat.Gif || format == OfficeImageFormat.Tiff || format == OfficeImageFormat.Webp;

    private static bool RequiresDownsampling(OfficeImageInfo info, int targetWidth, int targetHeight, double threshold) =>
        info.Width > targetWidth * threshold || info.Height > targetHeight * threshold;

    private static int ResolvePlacementPixelSize(double points, double dpi) {
        double pixels = Math.Ceiling(Math.Abs(points) * dpi / 72D);
        return pixels >= int.MaxValue ? int.MaxValue : Math.Max(1, (int)pixels);
    }

    private static string BuildPlacementOptimizationCacheKey(
        byte[] data,
        int targetWidth,
        int targetHeight,
        PdfImageOptimizationOptions options) {
        using var hash = SHA256.Create();
        AppendHashBytes(hash, data);
        hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        string sourceHash = ToHex(hash.Hash ?? Array.Empty<byte>());
        return sourceHash + ":" +
            ((int)options.Mode).ToString(CultureInfo.InvariantCulture) + ":" +
            ((int)options.MetadataPolicy).ToString(CultureInfo.InvariantCulture) + ":" +
            ((int)options.MetadataSelection).ToString(CultureInfo.InvariantCulture) + ":" +
            targetWidth.ToString(CultureInfo.InvariantCulture) + "x" +
            targetHeight.ToString(CultureInfo.InvariantCulture) + ":" +
            ((int)options.ResamplingMode).ToString(CultureInfo.InvariantCulture) + ":" +
            options.JpegQuality.ToString(CultureInfo.InvariantCulture) + ":" +
            (options.KeepOriginalWhenNotSmaller ? "1" : "0");
    }
}
