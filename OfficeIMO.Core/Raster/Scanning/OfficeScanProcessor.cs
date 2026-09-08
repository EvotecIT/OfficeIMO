using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Dependency-free, bounded preparation of document scans with reversible coordinate mapping.</summary>
public static partial class OfficeScanProcessor {
    /// <summary>Processes a separately owned image while preserving the source and reporting every requested transformation.</summary>
    /// <exception cref="OfficeScanProcessingLimitException">A configured pixel, buffer, or analysis-work limit cannot be met.</exception>
    public static OfficeScanProcessingResult Process(OfficeRasterImage source, OfficeScanProcessingOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        OfficeScanProcessingOptions effective = (options ?? new OfficeScanProcessingOptions()).Clone();
        effective.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        long sourcePixels = (long)source.Width * source.Height;
        CheckPixels(sourcePixels, effective);
        long analysisBytes = sourcePixels * (effective.Deskew && effective.NormalizeBackground ? 6L : 1L) +
            (long)effective.MaximumAnalysisSamples * 16L + 64_000L;
        CheckMemory(sourcePixels * 4L + analysisBytes, effective);
        var steps = new List<OfficeScanProcessingStep>();
        double quarterRotation = effective.ClockwiseQuarterTurns * 90D;
        steps.Add(new OfficeScanProcessingStep("orientation", effective.ClockwiseQuarterTurns != 0,
            "Explicit clockwise quarter-turns: " + effective.ClockwiseQuarterTurns.ToString(CultureInfo.InvariantCulture) + "."));
        SkewEstimate skew = effective.Deskew ? EstimateSkew(source, effective, cancellationToken) : default;
        bool correctSkew = effective.Deskew && skew.Confidence >= effective.MinimumDeskewConfidence &&
            Math.Abs(skew.Degrees) >= 0.1D && Math.Abs(skew.Degrees) < effective.MaximumDeskewAngleDegrees - 0.1D;
        double correction = correctSkew ? -skew.Degrees : 0D;
        steps.Add(new OfficeScanProcessingStep("deskew", correctSkew,
            !effective.Deskew ? "Deskew was disabled." : correctSkew ? "Applied the confident text-line skew correction." :
            "Retained source orientation: no confident in-range skew correction was found."));
        OfficeTransform rotation = OfficeTransform.RotateDegrees(quarterRotation + correction);
        var bounds = rotation.TransformRectangleBounds(0D, 0D, source.Width, source.Height);
        int width = checked((int)Math.Ceiling(bounds.Right - bounds.Left - 0.0000001D));
        int height = checked((int)Math.Ceiling(bounds.Bottom - bounds.Top - 0.0000001D));
        CheckPixels((long)width * height, effective);
        OfficeTransform transform = rotation.Then(OfficeTransform.Translate(-bounds.Left, -bounds.Top));
        // Includes retained source, transformation/output buffers, local background buffers, and analysis scratch.
        long workingBytes = sourcePixels * 4L + (long)width * height * 11L + analysisBytes;
        CheckMemory(workingBytes, effective);
        OfficeRasterImage output = OfficeRasterResampler.Transform(source, transform, width, height,
            OfficeColor.White, cancellationToken);
        ApplyPhotometricCorrections(output, effective, steps, cancellationToken);
        if (effective.MaximumDimension.HasValue && Math.Max(width, height) > effective.MaximumDimension.Value) {
            double ratio = effective.MaximumDimension.Value / (double)Math.Max(width, height);
            int targetWidth = Math.Max(1, (int)Math.Floor(width * ratio));
            int targetHeight = Math.Max(1, (int)Math.Floor(height * ratio));
            if (!OfficeRasterResampler.TryGetHighQualityWorkingSetBytes(width, height, targetWidth, targetHeight,
                    OfficeRasterResamplingMode.Area, sourcePixels * 4L, out long resizeBytes)) {
                throw new OfficeScanProcessingLimitException("Scan downsampling exceeds the managed resampling limit.");
            }
            workingBytes = Math.Max(workingBytes, resizeBytes);
            CheckMemory(workingBytes, effective);
            output = OfficeRasterResampler.Resize(output, targetWidth, targetHeight, OfficeRasterResamplingMode.Area,
                OfficeRasterResamplingColorSpace.EncodedSrgb, sourcePixels * 4L, cancellationToken);
            transform = transform.Then(OfficeTransform.Scale(targetWidth / (double)width, targetHeight / (double)height));
            width = targetWidth; height = targetHeight;
            // Area sampling creates gray edge samples; honor the requested final bilevel contract.
            if (effective.ColorMode == OfficeScanColorMode.Bilevel) ApplyBilevel(output, effective.BilevelThreshold, cancellationToken);
            steps.Add(new OfficeScanProcessingStep("downsample", true, "Reduced the longest side using area sampling; no upscaling was performed."));
        } else {
            steps.Add(new OfficeScanProcessingStep("downsample", false, "Retained pixel dimensions; no downsampling was required."));
        }
        (double foreground, bool blank) = MeasureForeground(output, cancellationToken);
        steps.Add(new OfficeScanProcessingStep("blank-page", false, blank ?
            "The image is probably blank. It has been retained for caller review." : "Foreground was detected. The page has been retained."));
        var report = new OfficeScanProcessingReport(source.Width, source.Height, width, height, transform,
            skew.Degrees, correction, skew.Confidence, foreground, blank, workingBytes, steps);
        return new OfficeScanProcessingResult(output, report);
    }

    private static void CheckPixels(long pixels, OfficeScanProcessingOptions options) {
        if (pixels <= 0 || pixels > options.MaximumPixels) throw new OfficeScanProcessingLimitException("Scan pixel count exceeds MaximumPixels.");
    }
    private static void CheckMemory(long bytes, OfficeScanProcessingOptions options) {
        if (bytes > options.MaximumWorkingBytes) throw new OfficeScanProcessingLimitException("Scan buffers exceed MaximumWorkingBytes.");
    }
}
