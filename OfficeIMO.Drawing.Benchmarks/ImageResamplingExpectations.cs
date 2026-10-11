namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Owns independently recorded output fingerprints for timed resampling workloads.</summary>
internal static class ImageResamplingExpectations {
    private static readonly IReadOnlyDictionary<(string ScenarioId, OfficeRasterResamplingMode Mode), string>
        ExpectedFingerprints = new Dictionary<(string, OfficeRasterResamplingMode), string> {
            // Photo uses the common decoder's high-quality JPEG chroma reconstruction.
            [(ImageBenchmarkScenarios.Photo.Id, OfficeRasterResamplingMode.Bilinear)] = "016D731C2B87EE64",
            [(ImageBenchmarkScenarios.Photo.Id, OfficeRasterResamplingMode.Area)] = "5081530057EEDF3F",
            [(ImageBenchmarkScenarios.Photo.Id, OfficeRasterResamplingMode.Lanczos3)] = "2887D793D00D4C62",
            [(ImageBenchmarkScenarios.Text.Id, OfficeRasterResamplingMode.Bilinear)] = "460790F6491118FB",
            [(ImageBenchmarkScenarios.Text.Id, OfficeRasterResamplingMode.Area)] = "7AD4E0A128104701",
            [(ImageBenchmarkScenarios.Text.Id, OfficeRasterResamplingMode.Lanczos3)] = "1455E1C99033CF42",
            [(ImageBenchmarkScenarios.LineArt.Id, OfficeRasterResamplingMode.Bilinear)] = "A02AF38D9D1361DF",
            [(ImageBenchmarkScenarios.LineArt.Id, OfficeRasterResamplingMode.Area)] = "B426BC00158EDEA1",
            [(ImageBenchmarkScenarios.LineArt.Id, OfficeRasterResamplingMode.Lanczos3)] = "FC62671C5CB1DCB4",
            [(ImageBenchmarkScenarios.AlphaGraphic.Id, OfficeRasterResamplingMode.Bilinear)] = "E1CA0A78512109FE",
            [(ImageBenchmarkScenarios.AlphaGraphic.Id, OfficeRasterResamplingMode.Area)] = "64A5D2AAFEFC68FA",
            // Shared sinc(x)*sinc(x/3) evaluation changes half-step RGB rounding by
            // one level versus the expanded formula; independently qualified alpha is unchanged.
            [(ImageBenchmarkScenarios.AlphaGraphic.Id, OfficeRasterResamplingMode.Lanczos3)] = "EB5A8C48833A4FD8"
        };

    internal static void Validate(
        string scenarioId,
        OfficeRasterResamplingMode mode,
        OfficeRasterImage image,
        int expectedWidth,
        int expectedHeight) {
        if (image.Width != expectedWidth || image.Height != expectedHeight) {
            throw new InvalidOperationException(
                $"{scenarioId} {mode} produced {image.Width}x{image.Height}; expected {expectedWidth}x{expectedHeight}.");
        }
        if (!ExpectedFingerprints.TryGetValue((scenarioId, mode), out string? expected)) {
            throw new InvalidOperationException($"{scenarioId} {mode} has no reviewed output fingerprint.");
        }
        string actual = ImageBenchmarkScenarios.Fingerprint(image);
        if (!string.Equals(actual, expected, StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                $"{scenarioId} {mode} output fingerprint {actual} did not match {expected}.");
        }
    }
}
