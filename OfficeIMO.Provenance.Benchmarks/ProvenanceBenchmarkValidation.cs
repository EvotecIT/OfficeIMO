using System.IO.Compression;

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
        ValidateExactOutput(fixture, output);
    }

    internal static OfficeProvenanceReport Inspect(ProvenanceBenchmarkFixture fixture) =>
        OfficeProvenanceInspector.Inspect(fixture.Asset, fixture.FileName);

    internal static OfficeProvenanceRemovalResult Remove(ProvenanceBenchmarkFixture fixture) =>
        OfficeProvenanceRemover.Remove(fixture.Asset, fixture.FileName);

    internal static void ValidateExactOutput(ProvenanceBenchmarkFixture fixture, byte[] output) {
        if (fixture.ExpectedOutput != null) {
            if (!output.AsSpan().SequenceEqual(fixture.ExpectedOutput)) {
                throw new InvalidOperationException($"{fixture.Format}/{fixture.Scale} removal output did not match the exact expected bytes.");
            }
            return;
        }
        if (fixture.Format != "ZIP" || fixture.ExpectedPreservedPayload == null) {
            throw new InvalidOperationException($"{fixture.Format}/{fixture.Scale} has no exact removal-output contract.");
        }
        using var stream = new MemoryStream(output, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        if (archive.Entries.Count != 1 || archive.Entries[0].FullName != "payload.bin") {
            throw new InvalidOperationException($"{fixture.Format}/{fixture.Scale} removal did not preserve exactly the expected ZIP entry.");
        }
        using Stream entry = archive.Entries[0].Open();
        using var content = new MemoryStream();
        entry.CopyTo(content);
        if (!content.ToArray().AsSpan().SequenceEqual(fixture.ExpectedPreservedPayload)) {
            throw new InvalidOperationException($"{fixture.Format}/{fixture.Scale} removal changed the preserved ZIP payload.");
        }
    }
}
