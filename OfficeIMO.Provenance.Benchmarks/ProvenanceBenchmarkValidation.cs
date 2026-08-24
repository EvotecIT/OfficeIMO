namespace OfficeIMO.Provenance.Benchmarks;

internal static class ProvenanceBenchmarkValidation {
    internal static void ValidateAll(bool writeSummary) {
        foreach (string format in ProvenanceBenchmarkCorpus.Formats) {
            foreach (string scale in ProvenanceBenchmarkCorpus.Scales) {
                ProvenanceBenchmarkFixture fixture = ProvenanceBenchmarkCorpus.Create(format, scale);
                Validate(fixture);
                if (writeSummary) {
                    Console.WriteLine(
                        $"{format,-5} {scale,-5} input {fixture.Asset.Length,10:N0} bytes | " +
                        $"output {fixture.ExpectedOutputBytes,10:N0} bytes");
                }
            }
        }
    }

    internal static void Validate(ProvenanceBenchmarkFixture fixture) {
        OfficeProvenanceReport report = Inspect(fixture);
        if (!report.HasC2paManifest || report.Evidence.Count != 1 || !report.Evidence[0].IsStructurallyValid) {
            throw new InvalidOperationException($"{fixture.Format}/{fixture.Scale} did not expose one valid C2PA carrier.");
        }

        OfficeProvenanceRemovalResult result = Remove(fixture);
        if (!result.WasChanged || result.Changes.Count != 1 || result.After.HasC2paManifest) {
            throw new InvalidOperationException($"{fixture.Format}/{fixture.Scale} did not remove exactly one C2PA carrier.");
        }
        byte[] output = result.ToArray();
        if (output.Length != fixture.ExpectedOutputBytes) {
            throw new InvalidOperationException(
                $"{fixture.Format}/{fixture.Scale} output was {output.Length} bytes, expected {fixture.ExpectedOutputBytes}.");
        }
    }

    internal static OfficeProvenanceReport Inspect(ProvenanceBenchmarkFixture fixture) =>
        OfficeProvenanceInspector.Inspect(fixture.Asset, fixture.FileName);

    internal static OfficeProvenanceRemovalResult Remove(ProvenanceBenchmarkFixture fixture) =>
        OfficeProvenanceRemover.Remove(fixture.Asset, fixture.FileName);
}
