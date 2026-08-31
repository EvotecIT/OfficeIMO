using System.Text.Json;
using OfficeIMO.Pdf;

namespace OfficeIMO.PdfQualityCorpus;

internal static class PdfQualityCorpusManifest {
    internal static QualityManifest Load(string path) {
        byte[] bytes = File.ReadAllBytes(path);
        QualityManifest manifest = JsonSerializer.Deserialize<QualityManifest>(bytes, QualityJson.Options)
            ?? throw new InvalidDataException("PDF quality corpus manifest is empty.");
        Validate(manifest);
        return manifest;
    }

    internal static string ResolveCasePath(string rootDirectory, QualityCase item) {
        if (Path.IsPathRooted(item.File)) throw new InvalidDataException("Corpus file paths must be relative: " + item.Id + ".");
        string root = Path.GetFullPath(rootDirectory);
        string candidate = Path.GetFullPath(Path.Combine(root, item.File));
        string relative = Path.GetRelativePath(root, candidate);
        if (relative == ".." || relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal) || Path.IsPathRooted(relative)) {
            throw new InvalidDataException("Corpus file escapes the configured root: " + item.Id + ".");
        }
        RejectReparsePoints(root, candidate, item.Id);
        return candidate;
    }

    private static void Validate(QualityManifest manifest) {
        if (manifest.Version <= 0) throw new InvalidDataException("Manifest version must be positive.");
        RequireHttps(manifest.Authority, "authority");
        if (manifest.Sources.Count == 0) throw new InvalidDataException("Manifest must declare at least one source.");
        if (manifest.Cases.Count == 0) throw new InvalidDataException("Manifest must declare at least one case.");

        var sourceIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (QualitySource source in manifest.Sources) {
            RequireIdentifier(source.Id, "source id");
            if (!sourceIds.Add(source.Id)) throw new InvalidDataException("Duplicate source id: " + source.Id + ".");
            RequireHttps(source.Repository, "source repository");
            if (source.Commit.Length != 40 || source.Commit.Any(character => !Uri.IsHexDigit(character))) {
                throw new InvalidDataException("Source commit must be a 40-character hexadecimal hash: " + source.Id + ".");
            }
            if (string.IsNullOrWhiteSpace(source.License)) throw new InvalidDataException("Source license is required: " + source.Id + ".");
        }

        var caseIds = new HashSet<string>(StringComparer.Ordinal);
        var files = new HashSet<string>(StringComparer.Ordinal);
        foreach (QualityCase item in manifest.Cases) {
            RequireIdentifier(item.Id, "case id");
            if (!caseIds.Add(item.Id)) throw new InvalidDataException("Duplicate case id: " + item.Id + ".");
            if (!sourceIds.Contains(item.Source)) throw new InvalidDataException("Unknown source id for case: " + item.Id + ".");
            if (string.IsNullOrWhiteSpace(item.File) || !files.Add(item.File)) throw new InvalidDataException("Case file must be unique and nonempty: " + item.Id + ".");
            if (string.IsNullOrWhiteSpace(item.SourcePath)) throw new InvalidDataException("Source path is required: " + item.Id + ".");
            if (item.Sha256.Length != 64 || item.Sha256.Any(character => !Uri.IsHexDigit(character))) {
                throw new InvalidDataException("Case SHA-256 must be a 64-character hexadecimal hash: " + item.Id + ".");
            }
            if (item.ByteLength < 0 || item.PageCount < 0 || item.MinimumTextCharacters < 0 || item.MinimumAttachments < 0 || item.MinimumLinks < 0 || item.MinimumAnnotations < 0) {
                throw new InvalidDataException("Case numeric expectations cannot be negative: " + item.Id + ".");
            }
            if (item.MinimumOptionalContentGroups < 0 ||
                item.MinimumFonts < 0 ||
                item.MinimumEmbeddedFonts < 0 ||
                item.MinimumSubsetFonts < 0 ||
                item.MaximumMissingToUnicodeFonts < 0) {
                throw new InvalidDataException("Case optional numeric expectations cannot be negative: " + item.Id + ".");
            }
            if (!Enum.TryParse<PdfMutationExecutionMode>(item.ExpectedMutationMode, ignoreCase: false, out _)) {
                throw new InvalidDataException("Case expected mutation mode is invalid: " + item.Id + ".");
            }
            RequireUniqueValues(item.ExpectedAnnotationActionTypes, "annotation action type", item.Id);
            RequireUniqueValues(item.ExpectedRepairCodes, "repair code", item.Id);
            RequireUniqueValues(item.ExpectedRenderDiagnosticCodes, "render diagnostic code", item.Id);
            if (item.Features.Count == 0) throw new InvalidDataException("Case features are required: " + item.Id + ".");
            RequireUniqueValues(item.Features, "feature", item.Id);
        }
    }

    private static void RequireUniqueValues(IReadOnlyList<string> values, string name, string caseId) {
        var unique = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 0; index < values.Count; index++) {
            if (string.IsNullOrWhiteSpace(values[index]) || !unique.Add(values[index])) {
                throw new InvalidDataException("Case " + name + " values must be unique and nonempty: " + caseId + ".");
            }
        }
    }

    private static void RequireIdentifier(string value, string name) {
        if (string.IsNullOrWhiteSpace(value) || value.Any(character => !(char.IsLetterOrDigit(character) || character == '-' || character == '_' || character == '.'))) {
            throw new InvalidDataException("Invalid " + name + ": " + value + ".");
        }
    }

    private static void RequireHttps(string value, string name) {
        if (!Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) || uri.Scheme != Uri.UriSchemeHttps) {
            throw new InvalidDataException("Manifest " + name + " must be an absolute HTTPS URI.");
        }
    }

    private static void RejectReparsePoints(string root, string file, string caseId) {
        string? current = file;
        while (current is not null && !string.Equals(current, root, StringComparison.OrdinalIgnoreCase)) {
            if (File.Exists(current) || Directory.Exists(current)) {
                if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0) {
                    throw new InvalidDataException("Corpus path contains a reparse point: " + caseId + ".");
                }
            }
            current = Path.GetDirectoryName(current);
        }
    }
}
