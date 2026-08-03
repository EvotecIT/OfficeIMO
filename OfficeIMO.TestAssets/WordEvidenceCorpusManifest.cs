using System.Security.Cryptography;
using System.Text.Json;

namespace OfficeIMO.TestAssets;

internal static class WordEvidenceCorpusManifestLoader {
    private const string ManifestRelativePath = "Word/EvidenceCorpus/corpus-manifest.json";
    private static readonly HashSet<string> RequiredFamilies = new(StringComparer.OrdinalIgnoreCase) {
        "review", "redline", "template", "mail-merge", "legacy-doc", "word-html", "rendering", "performance"
    };

    internal static WordEvidenceCorpusManifest Load() {
        string path = ResolveDocumentPath(ManifestRelativePath);
        return JsonSerializer.Deserialize<WordEvidenceCorpusManifest>(File.ReadAllText(path), new JsonSerializerOptions {
            PropertyNameCaseInsensitive = true
        }) ?? throw new InvalidDataException("The Word evidence corpus manifest is empty: " + path);
    }

    internal static IReadOnlyList<string> Validate(WordEvidenceCorpusManifest manifest) {
        var errors = new List<string>();
        if (manifest.SchemaVersion != 1) errors.Add("Unsupported Word evidence corpus schemaVersion.");
        if (!IsSafeRelativePath(manifest.Provenance) || !File.Exists(ResolveDocumentPath(manifest.Provenance))) {
            errors.Add("The Word evidence corpus provenance file is missing or unsafe.");
        }

        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var coveredFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (WordEvidenceCorpusArtifact artifact in manifest.Artifacts) {
            string label = string.IsNullOrWhiteSpace(artifact.Id) ? "<unnamed>" : artifact.Id;
            if (!ids.Add(artifact.Id)) errors.Add("Duplicate artifact id: " + artifact.Id);
            if (artifact.Families.Count == 0 || artifact.Families.Any(string.IsNullOrWhiteSpace)) errors.Add(label + ": families are required.");
            foreach (string family in artifact.Families) coveredFamilies.Add(family);
            if (string.IsNullOrWhiteSpace(artifact.Producer)) errors.Add(label + ": producer is required.");
            if (artifact.Oracles.Count == 0 || artifact.Oracles.Any(string.IsNullOrWhiteSpace)) errors.Add(label + ": oracles are required.");
            if (string.IsNullOrWhiteSpace(artifact.LossPolicy)) errors.Add(label + ": lossPolicy is required.");
            if (string.IsNullOrWhiteSpace(artifact.SourceTest)) errors.Add(label + ": sourceTest is required.");
            if (string.IsNullOrWhiteSpace(artifact.Contract)) errors.Add(label + ": contract is required.");
            if (artifact.Contract.IndexOf("native DOC authoring is supported", StringComparison.OrdinalIgnoreCase) >= 0) {
                errors.Add(label + ": unsupported native DOC authoring claim.");
            }

            if (string.IsNullOrWhiteSpace(artifact.Path)) {
                if (string.IsNullOrWhiteSpace(artifact.Generator)) errors.Add(label + ": path or generator is required.");
            } else if (!IsSafeRelativePath(artifact.Path!)) {
                errors.Add(label + ": artifact path is unsafe.");
            } else {
                string artifactPath = ResolveDocumentPath(artifact.Path!);
                if (!File.Exists(artifactPath)) {
                    errors.Add(label + ": artifact is missing: " + artifact.Path);
                } else if (!TryComputeSha256(artifactPath, artifact.HashMode, out string actualHash)) {
                    errors.Add(label + ": unsupported hashMode: " + artifact.HashMode + ".");
                } else if (!string.Equals(actualHash, artifact.Sha256, StringComparison.OrdinalIgnoreCase)) {
                    errors.Add(label + ": SHA-256 mismatch.");
                }
            }

            if (!string.IsNullOrWhiteSpace(artifact.ApprovedReport) &&
                (!IsSafeRelativePath(artifact.ApprovedReport!) || !File.Exists(ResolveDocumentPath(artifact.ApprovedReport!)))) {
                errors.Add(label + ": approved report is missing or unsafe.");
            }
            if (artifact.Families.Contains("legacy-doc", StringComparer.OrdinalIgnoreCase) &&
                (!string.Equals(artifact.LossPolicy, "guarded", StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(artifact.ApprovedReport))) {
                errors.Add(label + ": legacy DOC evidence requires guarded loss and an approved report.");
            }
        }

        foreach (string family in RequiredFamilies) {
            if (!coveredFamilies.Contains(family)) errors.Add("Missing required Word evidence family: " + family);
        }
        return errors;
    }

    internal static string ResolveArtifactPath(WordEvidenceCorpusArtifact artifact) =>
        ResolveDocumentPath(artifact.Path ?? throw new InvalidOperationException("Generated corpus entries do not have artifact paths."));

    private static string ResolveDocumentPath(string relativePath) =>
        Path.Combine(AppContext.BaseDirectory, "Documents", relativePath.Replace('/', Path.DirectorySeparatorChar));

    private static bool IsSafeRelativePath(string path) =>
        !string.IsNullOrWhiteSpace(path) && !Path.IsPathRooted(path) &&
        !path.Replace('\\', '/').Split('/').Any(segment => segment is "." or "..");

    private static bool TryComputeSha256(string path, string hashMode, out string hash) {
        byte[] content;
        if (string.Equals(hashMode, "raw", StringComparison.OrdinalIgnoreCase)) {
            content = File.ReadAllBytes(path);
        } else if (string.Equals(hashMode, "canonical-text", StringComparison.OrdinalIgnoreCase)) {
            string text = File.ReadAllText(path).Replace("\r\n", "\n").Replace("\r", "\n");
            content = System.Text.Encoding.UTF8.GetBytes(text);
        } else {
            hash = string.Empty;
            return false;
        }

        using SHA256 sha256 = SHA256.Create();
        byte[] hashBytes = sha256.ComputeHash(content);
        var builder = new System.Text.StringBuilder(hashBytes.Length * 2);
        foreach (byte value in hashBytes) builder.Append(value.ToString("x2"));
        hash = builder.ToString();
        return true;
    }
}

internal sealed class WordEvidenceCorpusManifest {
    public int SchemaVersion { get; set; }
    public string Provenance { get; set; } = string.Empty;
    public List<WordEvidenceCorpusArtifact> Artifacts { get; set; } = new();
}

internal sealed class WordEvidenceCorpusArtifact {
    public string Id { get; set; } = string.Empty;
    public List<string> Families { get; set; } = new();
    public string? Path { get; set; }
    public string? Sha256 { get; set; }
    public string HashMode { get; set; } = "raw";
    public string? ApprovedReport { get; set; }
    public string? Generator { get; set; }
    public string Producer { get; set; } = string.Empty;
    public List<string> Oracles { get; set; } = new();
    public string LossPolicy { get; set; } = string.Empty;
    public string SourceTest { get; set; } = string.Empty;
    public string Contract { get; set; } = string.Empty;
}
