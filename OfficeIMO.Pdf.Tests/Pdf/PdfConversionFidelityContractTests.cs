using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConversionFidelityContractTests {
    [Fact]
    public void ConverterCatalog_UsesEvidenceBackedFidelityClaims() {
        string manifestPath = GetManifestPath();
        string repositoryRoot = Directory.GetParent(Path.GetDirectoryName(manifestPath)!)!.FullName;
        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(manifestPath));
        JsonElement root = document.RootElement;
        JsonElement quality = root.GetProperty("qualityContract");
        var levels = new HashSet<string>(ReadValues(quality.GetProperty("fidelityLevels")), StringComparer.Ordinal);
        var statuses = new HashSet<string>(ReadValues(quality.GetProperty("fidelityStatuses")), StringComparer.Ordinal);
        var scenarioIds = new HashSet<string>(
            root.GetProperty("scenarios").EnumerateArray().Select(item => Require(item, "id")),
            StringComparer.Ordinal);
        Dictionary<string, JsonElement> referenceEvidence = root.GetProperty("referenceEvidence")
            .EnumerateArray()
            .ToDictionary(item => Require(item, "id"), item => item, StringComparer.Ordinal);

        Assert.Contains("exact", levels);
        Assert.Contains("supported-with-approximation", levels);
        Assert.Contains("rasterized", levels);
        Assert.Contains("omitted", levels);
        Assert.Contains("unsupported", levels);
        Assert.Contains("externally-verified", statuses);
        Assert.False(string.IsNullOrWhiteSpace(Require(quality, "premiumClaimRule")));
        Assert.False(string.IsNullOrWhiteSpace(Require(quality, "exactScopeRule")));

        foreach (JsonElement converter in root.GetProperty("converterCatalog").EnumerateArray()) {
            string id = Require(converter, "id");
            Assert.Contains(Require(converter, "fidelityStatus"), statuses);
            string referencePolicy = Require(converter, "referencePolicy");
            var converterScenarioIds = new HashSet<string>(ReadValues(converter.GetProperty("scenarioIds")), StringComparer.Ordinal);
            Assert.False(string.IsNullOrWhiteSpace(referencePolicy));

            if (referencePolicy.Contains("microsoft", StringComparison.Ordinal) ||
                referencePolicy.Contains("standards-corpus", StringComparison.Ordinal)) {
                string[] evidenceIds = ReadValues(converter.GetProperty("referenceEvidenceIds"));
                Assert.NotEmpty(evidenceIds);
                foreach (string evidenceId in evidenceIds) {
                    Assert.True(referenceEvidence.TryGetValue(evidenceId, out JsonElement evidence), "Converter " + id + " references unknown reference evidence " + evidenceId + ".");
                    string evidenceScenarioId = Require(evidence, "scenarioId");
                    Assert.Contains(evidenceScenarioId, scenarioIds);
                    Assert.Contains(evidenceScenarioId, converterScenarioIds);
                    AssertReferenceEvidence(repositoryRoot, evidence);
                }
            }

            if (!converter.TryGetProperty("capabilityClaims", out JsonElement claims)) continue;
            Assert.NotEmpty(claims.EnumerateArray());
            foreach (JsonElement claim in claims.EnumerateArray()) {
                Assert.False(string.IsNullOrWhiteSpace(Require(claim, "capability")));
                Assert.Contains(Require(claim, "level"), levels);
                string[] evidence = ReadValues(claim.GetProperty("evidenceScenarioIds"));
                Assert.NotEmpty(evidence);
                foreach (string scenarioId in evidence) {
                    Assert.True(scenarioIds.Contains(scenarioId), "Converter " + id + " references unknown evidence scenario " + scenarioId + ".");
                }
            }
        }
    }

    private static void AssertReferenceEvidence(string repositoryRoot, JsonElement evidence) {
        string kind = Require(evidence, "kind");
        string scenarioId = Require(evidence, "scenarioId");
        if (string.Equals(kind, "external-office-reference", StringComparison.Ordinal)) {
            Assert.False(string.IsNullOrWhiteSpace(Require(evidence, "producer")));
            string metadataPath = ResolveRepositoryPath(repositoryRoot, Require(evidence, "metadataPath"));
            Assert.True(File.Exists(metadataPath), "External reference metadata is missing: " + metadataPath + ".");
            using JsonDocument metadata = JsonDocument.Parse(File.ReadAllText(metadataPath));
            Assert.Contains(
                metadata.RootElement.GetProperty("scenarios").EnumerateArray(),
                item => string.Equals(Require(item, "id"), scenarioId, StringComparison.Ordinal));
            return;
        }

        if (string.Equals(kind, "standards-corpus", StringComparison.Ordinal)) {
            string sourcePath = ResolveRepositoryPath(repositoryRoot, Require(evidence, "sourcePath"));
            string baselineDirectory = ResolveRepositoryPath(repositoryRoot, Require(evidence, "baselineDirectory"));
            Assert.True(File.Exists(sourcePath), "Standards corpus source is missing: " + sourcePath + ".");
            Assert.True(Directory.Exists(baselineDirectory), "Standards corpus baseline directory is missing: " + baselineDirectory + ".");
            Assert.NotEmpty(Directory.GetFiles(baselineDirectory, "officeimo-pdf-native-html-*.png", SearchOption.TopDirectoryOnly));
            return;
        }

        throw new Xunit.Sdk.XunitException("Unknown reference evidence kind: " + kind + ".");
    }

    private static string ResolveRepositoryPath(string repositoryRoot, string relativePath) =>
        Path.Combine(repositoryRoot, relativePath.Replace('/', Path.DirectorySeparatorChar));

    private static string[] ReadValues(JsonElement array) =>
        array.EnumerateArray()
            .Select(item => item.GetString())
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value!)
            .ToArray();

    private static string Require(JsonElement element, string property) {
        string? value = element.GetProperty(property).GetString();
        Assert.False(string.IsNullOrWhiteSpace(value));
        return value!;
    }

    private static string GetManifestPath() {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            string candidate = Path.Combine(directory.FullName, "Docs", "pdf-conversion-scenarios.json");
            if (File.Exists(candidate)) return candidate;
            directory = directory.Parent;
        }

        throw new FileNotFoundException("Could not locate Docs/pdf-conversion-scenarios.json from test runtime base directory.");
    }
}
