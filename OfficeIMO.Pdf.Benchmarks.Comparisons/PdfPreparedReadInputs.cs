using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Prepares and admits optional identical bytes for the generated-reader matrix.</summary>
internal static class PdfPreparedReadInputs {
    private const string ManifestName = "pdf-read-inputs.json";
    private static readonly PdfBenchmarkScale[] Scales = [PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High];
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        WriteIndented = true
    };

    internal static int Prepare(string outputDirectory) {
        string root = Path.GetFullPath(outputDirectory);
        if (Directory.Exists(root) || File.Exists(root)) {
            throw new IOException("Use a fresh generated-reader input directory.");
        }
        Directory.CreateDirectory(root);
        var entries = new List<PreparedInput>();
        foreach (PdfBenchmarkScale scale in Scales) {
            PdfBenchmarkScenario scenario = PdfBenchmarkScenario.Get(scale);
            foreach (PdfBenchmarkProducer producer in Enum.GetValues<PdfBenchmarkProducer>()) {
                byte[] bytes = PdfDocumentGenerators.Generate(producer, scenario);
                PdfBenchmarkValidation.ValidateGenerated(bytes, scenario, producer.ToString());
                string fileName = FileName(producer, scale);
                File.WriteAllBytes(Path.Combine(root, fileName), bytes);
                entries.Add(new(producer.ToString(), scale.ToString(), fileName, bytes.Length, Hash(bytes)));
            }
        }
        Assembly assembly = typeof(PdfReadBenchmarks).Assembly;
        var manifest = new PreparedManifest(1,
            assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()!.InformationalVersion,
            Hash(File.ReadAllBytes(assembly.Location)), entries.ToArray());
        string path = Path.Combine(root, ManifestName);
        File.WriteAllText(path, JsonSerializer.Serialize(manifest, JsonOptions));
        Console.WriteLine($"PDF_READ_MANIFEST|{Hash(File.ReadAllBytes(path))}|{entries.Count}");
        return 0;
    }

    internal static byte[] Resolve(PdfBenchmarkProducer producer, PdfBenchmarkScenario scenario) {
        string? root = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_READ_INPUT_ROOT");
        if (string.IsNullOrWhiteSpace(root)) {
            return PdfDocumentGenerators.Generate(producer, scenario);
        }
        root = Path.GetFullPath(root);
        string? expectedManifestHash = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_READ_INPUT_MANIFEST_SHA256");
        byte[] manifestBytes = File.ReadAllBytes(Path.Combine(root, ManifestName));
        string manifestHash = Hash(manifestBytes);
        if (!IsSha256(expectedManifestHash) ||
            !manifestHash.Equals(expectedManifestHash, StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidDataException("The prepared-reader manifest does not match its required SHA-256.");
        }
        PreparedManifest manifest = JsonSerializer.Deserialize<PreparedManifest>(manifestBytes, JsonOptions)
            ?? throw new InvalidDataException("The prepared-reader manifest is empty.");
        PdfBenchmarkProducer[] producers = Enum.GetValues<PdfBenchmarkProducer>();
        if (manifest.SchemaVersion != 1 || manifest.Entries is null ||
            manifest.Entries.Length != Scales.Length * producers.Length ||
            Scales.Any(scale => producers.Any(engine => manifest.Entries.Count(entry =>
                entry is not null && entry.Scale == scale.ToString() && entry.Producer == engine.ToString()) != 1))) {
            throw new InvalidDataException("The prepared-reader manifest must identify every producer and scale exactly once.");
        }
        PreparedInput selected = manifest.Entries.Single(entry =>
            entry.Scale == scenario.Scale.ToString() && entry.Producer == producer.ToString());
        if (selected.FileName != FileName(producer, scenario.Scale) || selected.Bytes < 5 || !IsSha256(selected.Sha256)) {
            throw new InvalidDataException("Invalid prepared-reader input identity.");
        }
        byte[] bytes = File.ReadAllBytes(Path.Combine(root, selected.FileName));
        string actualHash = Hash(bytes);
        if (bytes.Length != selected.Bytes || !actualHash.Equals(selected.Sha256, StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidDataException($"Prepared-reader bytes differ for {producer}/{scenario.Scale}.");
        }
        Console.WriteLine($"PDF_READ_INPUT|{scenario.Scale}|{producer}|{actualHash}|{bytes.Length}|{manifestHash}");
        return bytes;
    }

    private static string FileName(PdfBenchmarkProducer producer, PdfBenchmarkScale scale) => $"{producer}-{scale}.pdf";
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));
    private static bool IsSha256(string? value) => value is { Length: 64 } && value.All(Uri.IsHexDigit);

    private sealed record PreparedManifest(int SchemaVersion, string SourceVersion, string BenchmarkSha256, PreparedInput[] Entries);
    private sealed record PreparedInput(string Producer, string Scale, string FileName, int Bytes, string Sha256);
}
