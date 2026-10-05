using System.Text.Json;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeProvenanceSignalMeasurementTests {
    [Fact]
    public async Task ProviderMeasurementsSurviveCanonicalReportSerialization() {
        var measurement = new OfficeProvenanceSignalMeasurement("1.2.3", "example-watermark", "z-score", 4.2,
            threshold: 3.5, tokenCount: 400, tokenizer: "example-tokenizer/1", configurationId: "public-profile-1",
            calibrationReference: "urn:example:calibration:1", textSha256: new string('A', 64), textOffset: 12, textLength: 800);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
        OfficeProvenanceWorkflowResult result;
        try {
            await File.WriteAllTextAsync(path, "sample");
            result = await new OfficeWorkflowRunner(null, [new MeasurementDetector(measurement)])
                .RunProvenanceAsync(new OfficeProvenanceWorkflowRequest { InputPath = path, Operation = OfficeProvenanceWorkflowOperation.Assess });
        } finally { File.Delete(path); }
        Assert.True(result.Succeeded, result.Summary);
        using JsonDocument document = JsonDocument.Parse(OfficeProvenanceReportSerializer.Serialize(result));
        JsonElement evidence = document.RootElement.GetProperty("assessment").GetProperty("providerSignals")[0].GetProperty("measurement");
        Assert.Equal(4.2, evidence.GetProperty("score").GetDouble());
        Assert.Equal("z-score", evidence.GetProperty("scoreName").GetString());
        Assert.Equal(new string('a', 64), evidence.GetProperty("textSha256").GetString());
        Assert.Equal(400, evidence.GetProperty("tokenCount").GetInt32());
        Assert.Equal(12, evidence.GetProperty("textOffset").GetInt32());
        Assert.Equal("urn:example:calibration:1", evidence.GetProperty("calibrationReference").GetString());
    }

    private sealed class MeasurementDetector(OfficeProvenanceSignalMeasurement measurement) : IOfficeProvenanceSignalDetector {
        public string Name => "test-provider";
        public OfficeProvenanceSignalKind SignalKind => OfficeProvenanceSignalKind.StatisticalTextWatermark;
        public OfficeProvenanceSignalResult Detect(string filePath) => new(Name, SignalKind, OfficeProvenanceSignalStatus.Detected, [], measurement);
    }

    [Fact]
    public void MeasurementsRejectUnserializableScoresAndUnboundSpans() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeProvenanceSignalMeasurement("1", "algorithm", "score", double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeProvenanceSignalMeasurement("1", "algorithm", "score", 1, threshold: double.PositiveInfinity));
        Assert.Throws<ArgumentException>(() => new OfficeProvenanceSignalMeasurement("1", "algorithm", "score", 1, textOffset: 0, textLength: 10));
        Assert.Throws<ArgumentException>(() => new OfficeProvenanceSignalMeasurement("1", "algorithm", "score", 1, textSha256: "invalid"));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeProvenanceSignalMeasurement("1", "algorithm", "score", 1,
            textSha256: new string('0', 64), textOffset: int.MaxValue, textLength: 1));
        Assert.Null(new OfficeProvenanceSignalResult("legacy", OfficeProvenanceSignalKind.StatisticalTextWatermark,
            OfficeProvenanceSignalStatus.Inconclusive).Measurement);
    }
}
