using System.Security.Cryptography;
using System.Text.Json;

namespace OfficeIMO.TestAssets;

internal static class OfficeInteroperabilityCorpusManifestLoader {
    private const string ManifestRelativePath = "OfficeInteroperabilityCorpus/corpus-manifest.json";
    private static readonly IReadOnlyDictionary<string, string> FormatIds =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            ["doc"] = "Word.Doc",
            ["xls"] = "Excel.Xls",
            ["xlsb"] = "Excel.Xlsb",
            ["ppt"] = "PowerPoint.Ppt"
        };
    private static readonly HashSet<string> Directions = new(StringComparer.OrdinalIgnoreCase) {
        "legacy-import",
        "new-legacy-write",
        "legacy-roundtrip",
        "modern-to-legacy",
        "legacy-to-modern"
    };
    private static readonly HashSet<string> Oracles = new(StringComparer.OrdinalIgnoreCase) {
        "structural",
        "semantic",
        "diagnostic",
        "visual",
        "microsoft-office-open",
        "libreoffice-open",
        "libreoffice-render"
    };

    internal static OfficeInteroperabilityCorpusManifest Load() {
        string manifestPath = Path.Combine(
            AppContext.BaseDirectory,
            "Documents",
            ManifestRelativePath.Replace('/', Path.DirectorySeparatorChar));
        string json = File.ReadAllText(manifestPath, Encoding.UTF8);
        return JsonSerializer.Deserialize<OfficeInteroperabilityCorpusManifest>(json, new JsonSerializerOptions {
            PropertyNameCaseInsensitive = true
        }) ?? throw new InvalidDataException($"The Office interoperability corpus manifest is empty: {manifestPath}");
    }

    internal static IReadOnlyList<string> Validate(OfficeInteroperabilityCorpusManifest manifest) {
        var errors = new List<string>();
        if (manifest.SchemaVersion != 2) {
            errors.Add($"Unsupported schemaVersion {manifest.SchemaVersion}; expected 2.");
        }
        if (manifest.Collections.Count == 0) {
            errors.Add("The manifest does not contain any collections.");
            return errors;
        }

        var collectionIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var artifactPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (OfficeInteroperabilityCorpusCollection collection in manifest.Collections) {
            ValidateCollection(collection, collectionIds, artifactPaths, errors);
        }

        return errors;
    }

    internal static string ResolveArtifactPath(
        OfficeInteroperabilityCorpusCollection collection,
        OfficeInteroperabilityCorpusArtifact artifact) {
        return ResolveDocumentPath(CombineRelative(collection.Root, artifact.File));
    }

    private static void ValidateCollection(
        OfficeInteroperabilityCorpusCollection collection,
        ISet<string> collectionIds,
        ISet<string> artifactPaths,
        ICollection<string> errors) {
        string label = string.IsNullOrWhiteSpace(collection.Id) ? "<unnamed collection>" : collection.Id;
        if (!collectionIds.Add(collection.Id)) {
            errors.Add($"Duplicate collection id: {collection.Id}");
        }
        if (!IsSafeRelativePath(collection.Root)) {
            errors.Add($"{label}: root must be a safe relative path: {collection.Root}");
            return;
        }
        if (!FormatIds.TryGetValue(collection.Format, out string? expectedFormatId)) {
            errors.Add($"{label}: unsupported format '{collection.Format}'.");
        } else if (!string.Equals(collection.FormatId, expectedFormatId, StringComparison.Ordinal)) {
            errors.Add($"{label}: formatId must be '{expectedFormatId}' for format '{collection.Format}'.");
        }
        if (collection.Role is not ("compatibility" or "diagnostic")) {
            errors.Add($"{label}: unsupported role '{collection.Role}'.");
        }
        if (string.IsNullOrWhiteSpace(collection.Producer)) {
            errors.Add($"{label}: producer is required.");
        }
        ValidateStringSet(label, "direction", collection.Directions, Directions, errors);
        ValidateStringSet(label, "oracle", collection.Oracles, Oracles, errors);
        if (!IsSafeRelativePath(collection.Provenance)) {
            errors.Add($"{label}: provenance must be a safe relative path: {collection.Provenance}");
        } else if (!File.Exists(ResolveDocumentPath(CombineRelative(collection.Root, collection.Provenance)))) {
            errors.Add($"{label}: provenance file does not exist: {collection.Provenance}");
        }
        if (collection.Artifacts.Count == 0) {
            errors.Add($"{label}: no artifacts are declared.");
            return;
        }

        var declared = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (OfficeInteroperabilityCorpusArtifact artifact in collection.Artifacts) {
            ValidateArtifact(collection, artifact, declared, artifactPaths, errors);
        }

        string rootPath = ResolveDocumentPath(collection.Root);
        if (!Directory.Exists(rootPath)) {
            errors.Add($"{label}: collection root does not exist: {collection.Root}");
            return;
        }

        string[] actual = Directory.GetFiles(rootPath, "*." + collection.Format, SearchOption.AllDirectories)
            .Select(path => NormalizeRelativePath(GetRelativePath(rootPath, path)))
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        string[] missingFromManifest = actual.Where(path => !declared.Contains(path)).ToArray();
        string[] missingFromDisk = declared.Where(path => !actual.Contains(path, StringComparer.OrdinalIgnoreCase)).ToArray();
        if (missingFromManifest.Length > 0) {
            errors.Add($"{label}: untracked {collection.Format} artifacts: {string.Join(", ", missingFromManifest)}");
        }
        if (missingFromDisk.Length > 0) {
            errors.Add($"{label}: declared artifacts missing from disk: {string.Join(", ", missingFromDisk)}");
        }
    }

    private static void ValidateStringSet(
        string label,
        string kind,
        IReadOnlyList<string> values,
        ISet<string> allowed,
        ICollection<string> errors) {
        if (values.Count == 0) {
            errors.Add($"{label}: at least one {kind} is required.");
            return;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string value in values) {
            if (string.IsNullOrWhiteSpace(value)) {
                errors.Add($"{label}: {kind} values cannot be empty.");
            } else if (!allowed.Contains(value)) {
                errors.Add($"{label}: unsupported {kind} '{value}'.");
            } else if (!seen.Add(value)) {
                errors.Add($"{label}: duplicate {kind} '{value}'.");
            }
        }
    }

    private static void ValidateArtifact(
        OfficeInteroperabilityCorpusCollection collection,
        OfficeInteroperabilityCorpusArtifact artifact,
        ISet<string> declared,
        ISet<string> artifactPaths,
        ICollection<string> errors) {
        string label = $"{collection.Id}/{artifact.File}";
        if (!IsSafeRelativePath(artifact.File)) {
            errors.Add($"{label}: file must be a safe relative path.");
            return;
        }
        string normalizedFile = NormalizeRelativePath(artifact.File);
        if (!declared.Add(normalizedFile)) {
            errors.Add($"{label}: artifact is declared more than once in the collection.");
        }
        string relativePath = CombineRelative(collection.Root, normalizedFile);
        if (!artifactPaths.Add(relativePath)) {
            errors.Add($"{label}: artifact path is declared by more than one collection.");
        }
        if (!string.Equals(Path.GetExtension(normalizedFile), "." + collection.Format, StringComparison.OrdinalIgnoreCase)) {
            errors.Add($"{label}: extension does not match collection format '{collection.Format}'.");
        }
        if (artifact.Focus.Count == 0 || artifact.Focus.Any(string.IsNullOrWhiteSpace)) {
            errors.Add($"{label}: at least one non-empty focus contract is required.");
        }

        string artifactPath = ResolveDocumentPath(relativePath);
        if (!File.Exists(artifactPath)) {
            errors.Add($"{label}: artifact does not exist.");
            return;
        }
        string actualHash = ComputeSha256(artifactPath);
        if (!string.Equals(actualHash, artifact.Sha256, StringComparison.OrdinalIgnoreCase)) {
            errors.Add($"{label}: SHA-256 mismatch. Expected {artifact.Sha256}, got {actualHash}.");
        }

        if (string.IsNullOrWhiteSpace(artifact.ApprovedReport)) {
            if (collection.Format is "doc" or "xls") {
                errors.Add($"{label}: legacy binary artifacts require an approved import report.");
            }
        } else if (!IsSafeRelativePath(artifact.ApprovedReport)) {
            errors.Add($"{label}: approvedReport must be a safe relative path.");
        } else if (!File.Exists(ResolveDocumentPath(CombineRelative(collection.Root, artifact.ApprovedReport)))) {
            errors.Add($"{label}: approved report does not exist: {artifact.ApprovedReport}");
        }
    }

    private static string ResolveDocumentPath(string relativePath) {
        return Path.Combine(
            AppContext.BaseDirectory,
            "Documents",
            relativePath.Replace('/', Path.DirectorySeparatorChar));
    }

    private static string CombineRelative(string left, string right) {
        return NormalizeRelativePath(left.TrimEnd('/') + "/" + right.TrimStart('/'));
    }

    private static string NormalizeRelativePath(string path) {
        return path.Replace('\\', '/').TrimStart('/');
    }

    private static bool IsSafeRelativePath(string path) {
        if (string.IsNullOrWhiteSpace(path) || Path.IsPathRooted(path)) {
            return false;
        }
        return !NormalizeRelativePath(path)
            .Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries)
            .Any(segment => segment is "." or "..");
    }

    private static string ComputeSha256(string path) {
        using SHA256 sha256 = SHA256.Create();
        using FileStream stream = File.OpenRead(path);
        byte[] hash = sha256.ComputeHash(stream);
        var text = new StringBuilder(hash.Length * 2);
        foreach (byte value in hash) {
            text.Append(value.ToString("x2"));
        }
        return text.ToString();
    }

    private static string GetRelativePath(string relativeTo, string path) {
#if NETFRAMEWORK
        Uri baseUri = new Uri(AppendDirectorySeparator(relativeTo));
        Uri pathUri = new Uri(path);
        return Uri.UnescapeDataString(baseUri.MakeRelativeUri(pathUri).ToString())
            .Replace('/', Path.DirectorySeparatorChar);
#else
        return Path.GetRelativePath(relativeTo, path);
#endif
    }

    private static string AppendDirectorySeparator(string path) {
        if (path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            || path.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)) {
            return path;
        }
        return path + Path.DirectorySeparatorChar;
    }
}

internal sealed class OfficeInteroperabilityCorpusManifest {
    public int SchemaVersion { get; set; }
    public List<OfficeInteroperabilityCorpusCollection> Collections { get; set; } = new();
}

internal sealed class OfficeInteroperabilityCorpusCollection {
    public string Id { get; set; } = string.Empty;
    public string Format { get; set; } = string.Empty;
    public string FormatId { get; set; } = string.Empty;
    public string Role { get; set; } = string.Empty;
    public string Root { get; set; } = string.Empty;
    public string Producer { get; set; } = string.Empty;
    public string Provenance { get; set; } = string.Empty;
    public List<string> Directions { get; set; } = new();
    public List<string> Oracles { get; set; } = new();
    public List<OfficeInteroperabilityCorpusArtifact> Artifacts { get; set; } = new();
}

internal sealed class OfficeInteroperabilityCorpusArtifact {
    public string File { get; set; } = string.Empty;
    public string Sha256 { get; set; } = string.Empty;
    public string? ApprovedReport { get; set; }
    public List<string> Focus { get; set; } = new();
}
