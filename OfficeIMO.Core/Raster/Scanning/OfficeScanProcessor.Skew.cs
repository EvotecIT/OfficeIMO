using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeScanProcessor {
    private readonly struct SkewEstimate {
        internal SkewEstimate(double degrees, double confidence) { Degrees = degrees; Confidence = confidence; }
        internal double Degrees { get; }
        internal double Confidence { get; }
    }

    private static SkewEstimate EstimateSkew(OfficeRasterImage image, OfficeScanProcessingOptions options, CancellationToken token) {
        // Shadows must not dominate the foreground histogram used to measure text-line alignment.
        byte[]? background = options.NormalizeBackground ? EstimateBackground(image, options.BackgroundRadius, token) : null;
        int stride = Math.Max(1, (int)Math.Ceiling(Math.Max(image.Width, image.Height) / 1000D));
        int width = (image.Width + stride - 1) / stride, height = (image.Height + stride - 1) / stride;
        var histogram = new int[256];
        int count = 0;
        for (int y = 0; y < image.Height; y += stride) {
            token.ThrowIfCancellationRequested();
            for (int x = 0; x < image.Width; x += stride) { histogram[Sample(x, y)]++; count++; }
        }
        int threshold = HistogramThreshold(histogram, count);
        int foreground = 0;
        for (int value = 0; value <= threshold; value++) foreground += histogram[value];
        if (foreground < 64 || foreground > count * 0.65D) return default;
        int takeEvery = Math.Max(1, (foreground + options.MaximumAnalysisSamples - 1) / options.MaximumAnalysisSamples);
        var points = new OfficePoint[Math.Min(foreground, options.MaximumAnalysisSamples)];
        int used = 0, seen = 0;
        for (int y = 0; y < image.Height; y += stride) {
            token.ThrowIfCancellationRequested();
            for (int x = 0; x < image.Width; x += stride) {
                if (Sample(x, y) > threshold) continue;
                if (seen++ % takeEvery != 0) continue;
                double px = x / (double)stride - width / 2D, py = y / (double)stride - height / 2D;
                points[used++] = options.ClockwiseQuarterTurns switch {
                    1 => new OfficePoint(-py, px), 2 => new OfficePoint(-px, -py),
                    3 => new OfficePoint(py, -px), _ => new OfficePoint(px, py)
                };
            }
        }
        if (used < 64) return default;
        var bins = new int[2 * (width + height) + 5];
        int center = bins.Length / 2;
        long operations = 0;
        double bestAngle = 0D, bestScore = 0D;
        int candidates = (int)Math.Ceiling(options.MaximumDeskewAngleDegrees * 8D) + 1;
        var scores = new double[candidates];
        for (int i = 0; i < candidates; i++) {
            double angle = -options.MaximumDeskewAngleDegrees + i * 0.25D;
            if (angle > options.MaximumDeskewAngleDegrees) break;
            double score = Score(angle);
            scores[i] = score;
            if (score > bestScore) { bestScore = score; bestAngle = angle; }
        }
        if (bestScore <= 0D) return default;
        double competingScore = 0D;
        for (int i = 0; i < candidates; i++) {
            double angle = -options.MaximumDeskewAngleDegrees + i * 0.25D;
            if (Math.Abs(angle - bestAngle) >= 0.75D) competingScore = Math.Max(competingScore, scores[i]);
        }
        double coarse = bestAngle;
        for (int i = -4; i <= 4; i++) {
            double angle = coarse + i * 0.05D;
            if (Math.Abs(angle) > options.MaximumDeskewAngleDegrees) continue;
            double score = Score(angle);
            if (score > bestScore) { bestScore = score; bestAngle = angle; }
        }
        return new SkewEstimate(bestAngle, Math.Max(0D, Math.Min(1D, (bestScore - competingScore) / bestScore)));

        int Sample(int x, int y) {
            int index = y * image.Width + x;
            int gray = Luminance(image.PixelBuffer, index * 4);
            int paper = background == null ? 255 : background[index];
            return paper <= 0 ? gray : Math.Min(255, (gray * 255 + paper / 2) / paper);
        }

        double Score(double angle) {
            token.ThrowIfCancellationRequested();
            if (used > options.MaximumAnalysisOperations - operations)
                throw new OfficeScanProcessingLimitException("Deskew analysis exceeds MaximumAnalysisOperations.");
            operations += used;
            Array.Clear(bins, 0, bins.Length);
            double radians = angle * Math.PI / 180D, sine = Math.Sin(radians), cosine = Math.Cos(radians);
            for (int i = 0; i < used; i++) {
                if ((i & 1023) == 0) token.ThrowIfCancellationRequested();
                int row = center + (int)Math.Floor(points[i].Y * cosine - points[i].X * sine + 0.5D);
                bins[row]++;
            }
            double score = 0D;
            for (int i = 0; i < bins.Length; i++) score += (double)bins[i] * bins[i];
            return score;
        }
    }
}
