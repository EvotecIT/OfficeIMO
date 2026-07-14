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
        Dictionary<string, OfficeImageOptimizationResult> cache) {
        PdfImageOptimizationOptions? options = documentOptions.ImageOptimizationSnapshot;
        if (options?.Enabled != true || !CanOptimizeImageFormat(image.Info.Format)) return;

        int targetWidth = ResolvePlacementPixelSize(image.W, options.TargetDpi);
        int targetHeight = ResolvePlacementPixelSize(image.H, options.TargetDpi);
        if (!RequiresDownsampling(image.Info, targetWidth, targetHeight, options.DownsampleThreshold)) return;

        string cacheKey = BuildPlacementOptimizationCacheKey(image.Data, targetWidth, targetHeight, options);
        if (!cache.TryGetValue(cacheKey, out OfficeImageOptimizationResult? result)) {
            try {
                result = OfficeImageOptimizer.Optimize(
                    image.Data,
                    new OfficeImageOptimizationRequest(targetWidth, targetHeight) {
                        ResamplingMode = options.ResamplingMode,
                        JpegQuality = options.JpegQuality,
                        KeepOriginalWhenNotSmaller = options.KeepOriginalWhenNotSmaller
                    });
            } catch (Exception exception) when (
                exception is ArgumentException ||
                exception is FormatException ||
                exception is InvalidOperationException ||
                exception is OverflowException) {
                return;
            }

            cache[cacheKey] = result;
        }

        if (!result.Changed) return;
        image.Data = result.Bytes;
        image.Info = result.Final;
    }

    private static bool CanOptimizeImageFormat(OfficeImageFormat format) =>
        format == OfficeImageFormat.Png || format == OfficeImageFormat.Jpeg || format == OfficeImageFormat.Bmp;

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
            targetWidth.ToString(CultureInfo.InvariantCulture) + "x" +
            targetHeight.ToString(CultureInfo.InvariantCulture) + ":" +
            ((int)options.ResamplingMode).ToString(CultureInfo.InvariantCulture) + ":" +
            options.JpegQuality.ToString(CultureInfo.InvariantCulture) + ":" +
            (options.KeepOriginalWhenNotSmaller ? "1" : "0");
    }
}
