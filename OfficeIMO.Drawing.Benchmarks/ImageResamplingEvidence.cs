namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Reports deterministic visual differences against pixel-area downsampling.</summary>
internal static class ImageResamplingEvidence {
    private sealed record EvidenceMode(
        string Label,
        OfficeRasterResamplingMode Mode,
        OfficeRasterResamplingColorSpace ColorSpace);

    private static readonly EvidenceMode[] Modes = {
        new("Bilinear", OfficeRasterResamplingMode.Bilinear, OfficeRasterResamplingColorSpace.EncodedSrgb),
        new("Area", OfficeRasterResamplingMode.Area, OfficeRasterResamplingColorSpace.EncodedSrgb),
        new("Lanczos3", OfficeRasterResamplingMode.Lanczos3, OfficeRasterResamplingColorSpace.EncodedSrgb),
        new("AreaLinear", OfficeRasterResamplingMode.Area, OfficeRasterResamplingColorSpace.LinearLight),
        new("LanczosLinear", OfficeRasterResamplingMode.Lanczos3, OfficeRasterResamplingColorSpace.LinearLight)
    };

    internal static void Validate(TextWriter writer) {
        writer.WriteLine();
        writer.WriteLine("Resampling fidelity matrix (pixel-area result is the antialiasing reference):");
        writer.WriteLine("Fixture        Mode          Output    Premul MAE     PSNR  Alpha MAE  Fingerprint");
        foreach (string scenarioId in ImageBenchmarkScenarios.ResamplingIds) {
            ImageBenchmarkScenario scenario = ImageBenchmarkScenarios.Get(scenarioId);
            OfficeRasterImage source = scenario.CreateImage();
            int width = Math.Max(1, source.Width / 4);
            int height = Math.Max(1, source.Height / 4);
            OfficeRasterImage reference = OfficeRasterResampler.Resize(
                source,
                width,
                height,
                OfficeRasterResamplingMode.Area);
            byte[] expected = reference.GetPixels();
            OfficeRasterImage linearReference = OfficeRasterResampler.Resize(
                source, width, height, OfficeRasterResamplingMode.Area,
                OfficeRasterResamplingColorSpace.LinearLight);
            foreach (EvidenceMode mode in Modes) {
                OfficeRasterImage actual = OfficeRasterResampler.Resize(source, width, height, mode.Mode, mode.ColorSpace);
                if (mode.ColorSpace == OfficeRasterResamplingColorSpace.EncodedSrgb) {
                    ImageResamplingExpectations.Validate(scenarioId, mode.Mode, actual, width, height);
                }
                byte[] pixels = actual.GetPixels();
                OfficeRasterImage repeated = OfficeRasterResampler.Resize(source, width, height, mode.Mode, mode.ColorSpace);
                if (!pixels.AsSpan().SequenceEqual(repeated.GetPixels())) {
                    throw new InvalidOperationException($"{scenarioId} {mode.Label} resampling was not deterministic.");
                }
                byte[] comparison = mode.ColorSpace == OfficeRasterResamplingColorSpace.LinearLight
                    ? linearReference.GetPixels()
                    : expected;
                (double mae, double psnr) = MeasurePremultipliedRgbFidelity(comparison, pixels);
                double alphaMae = MeasureAlphaError(comparison, pixels);
                writer.WriteLine(
                    $"{scenarioId,-14} {mode.Label,-13} {width,4}x{height,-4} " +
                    $"{mae,9:F3} {FormatPsnr(psnr),8} {alphaMae,10:F3}  {ImageBenchmarkScenarios.Fingerprint(actual)}");
            }
        }
    }

    internal static void WritePreviews(string outputDirectory, TextWriter writer) {
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            throw new ArgumentException("A preview output directory is required.", nameof(outputDirectory));
        }
        string fullPath = Path.GetFullPath(outputDirectory);
        Directory.CreateDirectory(fullPath);
        foreach (string scenarioId in ImageBenchmarkScenarios.ResamplingIds) {
            ImageBenchmarkScenario scenario = ImageBenchmarkScenarios.Get(scenarioId);
            OfficeRasterImage source = scenario.CreateImage();
            int width = Math.Max(1, source.Width / 4);
            int height = Math.Max(1, source.Height / 4);
            foreach (EvidenceMode mode in Modes) {
                OfficeRasterImage image = OfficeRasterResampler.Resize(source, width, height, mode.Mode, mode.ColorSpace);
                string path = Path.Combine(fullPath, $"{scenarioId}-{mode.Label}.png");
                File.WriteAllBytes(path, OfficePngWriter.Encode(image));
                writer.WriteLine(path);
            }
        }
    }

    private static double MeasureAlphaError(byte[] expected, byte[] actual) {
        if (expected.Length != actual.Length) {
            throw new InvalidOperationException("Resampling evidence dimensions did not match the reference.");
        }
        long error = 0L;
        int pixels = expected.Length / 4;
        for (int offset = 3; offset < expected.Length; offset += 4) {
            error += Math.Abs(expected[offset] - actual[offset]);
        }
        return error / (double)pixels;
    }

    private static (double MeanAbsoluteError, double Psnr) MeasurePremultipliedRgbFidelity(
        byte[] expected,
        byte[] actual) {
        if (expected.Length != actual.Length || expected.Length % 4 != 0) {
            throw new InvalidOperationException("Resampling evidence dimensions did not match the reference.");
        }
        double absoluteError = 0D;
        double squaredError = 0D;
        long channelCount = expected.LongLength / 4L * 3L;
        for (int offset = 0; offset < expected.Length; offset += 4) {
            double expectedAlpha = expected[offset + 3] / 255D;
            double actualAlpha = actual[offset + 3] / 255D;
            for (int channel = 0; channel < 3; channel++) {
                double difference = expected[offset + channel] * expectedAlpha -
                    actual[offset + channel] * actualAlpha;
                absoluteError += Math.Abs(difference);
                squaredError += difference * difference;
            }
        }
        double meanAbsoluteError = absoluteError / channelCount;
        double meanSquaredError = squaredError / channelCount;
        double psnr = meanSquaredError == 0D
            ? double.PositiveInfinity
            : 10D * Math.Log10(255D * 255D / meanSquaredError);
        return (meanAbsoluteError, psnr);
    }

    private static string FormatPsnr(double psnr) =>
        double.IsPositiveInfinity(psnr) ? "exact" : psnr.ToString("F2");
}
