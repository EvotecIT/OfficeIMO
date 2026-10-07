namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Owns independently recorded output fingerprints for timed resampling workloads.</summary>
internal static class ImageResamplingExpectations {
    private static readonly IReadOnlyDictionary<(string ScenarioId, OfficeRasterResamplingMode Mode), string>
        ExpectedFingerprints = new Dictionary<(string, OfficeRasterResamplingMode), string> {
            // Photo uses the qualified JPEG decoder's combined green rounding.
            [(ImageBenchmarkScenarios.Photo.Id, OfficeRasterResamplingMode.Bilinear)] = "AC957B343C6BCC7F",
            [(ImageBenchmarkScenarios.Photo.Id, OfficeRasterResamplingMode.Area)] = "ACF073D59C89EFCA",
            [(ImageBenchmarkScenarios.Photo.Id, OfficeRasterResamplingMode.Lanczos3)] = "58AC0410839A81FE",
            [(ImageBenchmarkScenarios.Text.Id, OfficeRasterResamplingMode.Bilinear)] = "460790F6491118FB",
            [(ImageBenchmarkScenarios.Text.Id, OfficeRasterResamplingMode.Area)] = "7AD4E0A128104701",
            [(ImageBenchmarkScenarios.Text.Id, OfficeRasterResamplingMode.Lanczos3)] = "1455E1C99033CF42",
            [(ImageBenchmarkScenarios.LineArt.Id, OfficeRasterResamplingMode.Bilinear)] = "A02AF38D9D1361DF",
            [(ImageBenchmarkScenarios.LineArt.Id, OfficeRasterResamplingMode.Area)] = "B426BC00158EDEA1",
            [(ImageBenchmarkScenarios.LineArt.Id, OfficeRasterResamplingMode.Lanczos3)] = "FC62671C5CB1DCB4",
            [(ImageBenchmarkScenarios.AlphaGraphic.Id, OfficeRasterResamplingMode.Bilinear)] = "E1CA0A78512109FE",
            [(ImageBenchmarkScenarios.AlphaGraphic.Id, OfficeRasterResamplingMode.Area)] = "64A5D2AAFEFC68FA",
            [(ImageBenchmarkScenarios.AlphaGraphic.Id, OfficeRasterResamplingMode.Lanczos3)] = "CE8EF8FB3D082605"
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
