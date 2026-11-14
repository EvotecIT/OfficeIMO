using System.Text;
using System.Text.Json;

namespace OfficeIMO.PdfCorpusReliability;

/// <summary>
/// Builds the file list from KillerPDF-style metadata (manifests, published test passwords and qpdf baselines),
/// or from a plain directory scan when no metadata checkout is supplied.
/// </summary>
internal static class CorpusInventory {
    public static List<CorpusEntry> Load(RunOptions options, out int missing) {
        missing = 0;
        List<CorpusEntry> entries = options.MetadataRoot is null
            ? ScanDirectory(options.CorpusRoot)
            : FromManifests(options, ref missing);
        if (options.Only.Count > 0) {
            entries = entries.Where(entry => options.Only.Any(only =>
                string.Equals(Normalize(entry.Path), Normalize(only), StringComparison.OrdinalIgnoreCase)
                || string.Equals(entry.Sha256, only, StringComparison.OrdinalIgnoreCase))).ToList();
        }
        if (options.Limit is int limit) entries = entries.Take(limit).ToList();
        return entries;
    }

    private static List<CorpusEntry> ScanDirectory(string root) =>
        Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories)
            .Where(static path => path.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            .OrderBy(static path => path, StringComparer.Ordinal)
            .Select(path => new CorpusEntry {
                Path = Normalize(Path.GetRelativePath(root, path)),
                FullPath = path,
                Collection = Normalize(Path.GetRelativePath(root, path)).Split('/')[0],
                Bytes = new FileInfo(path).Length
            })
            .ToList();

    private static List<CorpusEntry> FromManifests(RunOptions options, ref int missing) {
        string manifests = Path.Combine(options.MetadataRoot!, "manifests");
        Dictionary<string, string[]?> passwords = ReadPasswords(Path.Combine(manifests, $"v{options.CorpusVersion}-passwords.json"));
        HashSet<string> certificates = ReadCertificateKeys(Path.Combine(manifests, $"v{options.CorpusVersion}-certificates.json"));
        Dictionary<string, int> qpdf = ReadQpdfBaselines(Path.Combine(options.MetadataRoot!, "baselines"));

        var entries = new List<CorpusEntry>();
        foreach (string collection in options.Collections) {
            string manifest = Path.Combine(manifests, $"v{options.CorpusVersion}-{collection}.csv");
            if (!File.Exists(manifest)) throw new FileNotFoundException($"Manifest for collection '{collection}' was not found.", manifest);
            foreach (Dictionary<string, string> row in ReadCsv(manifest)) {
                string path = Normalize(row["Path"]);
                string? full = Resolve(options.CorpusRoot, collection, path);
                if (full is null) {
                    missing++;
                    continue;
                }
                string sha = row["SHA256"].ToUpperInvariant();
                passwords.TryGetValue(path, out string[]? candidates);
                entries.Add(new CorpusEntry {
                    Path = path,
                    FullPath = full,
                    Collection = collection,
                    Source = row.TryGetValue("Source", out string? source) ? source : "",
                    Sha256 = sha,
                    Bytes = long.TryParse(row["Bytes"], out long bytes) ? bytes : 0,
                    QpdfExitCode = qpdf.TryGetValue(sha, out int code) ? code : null,
                    Passwords = candidates,
                    PasswordsUnpublished = passwords.ContainsKey(path) && candidates is null,
                    CertificateEncrypted = certificates.Contains(path)
                });
            }
        }
        return entries;
    }

    private static string? Resolve(string root, string collection, string path) {
        string[] candidates = {
            Path.Combine(root, collection, path),
            Path.Combine(root, path)
        };
        return candidates.FirstOrDefault(File.Exists);
    }

    private static Dictionary<string, string[]?> ReadPasswords(string path) {
        var result = new Dictionary<string, string[]?>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(path)) return result;
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(path));
        foreach (JsonProperty property in json.RootElement.EnumerateObject()) {
            result[Normalize(property.Name)] = property.Value.ValueKind == JsonValueKind.Array
                ? property.Value.EnumerateArray().Select(static value => value.GetString() ?? "").ToArray()
                : null;
        }
        return result;
    }

    private static HashSet<string> ReadCertificateKeys(string path) {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(path)) return result;
        using JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(path));
        foreach (JsonProperty property in json.RootElement.EnumerateObject()) result.Add(Normalize(property.Name));
        return result;
    }

    /// <summary>Maps SHA-256 to the qpdf --check exit code recorded by the corpus (0 valid, 3 warnings, 2 errors).</summary>
    private static Dictionary<string, int> ReadQpdfBaselines(string directory) {
        var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        if (!Directory.Exists(directory)) return result;
        foreach (string file in Directory.EnumerateFiles(directory, "*qpdf*.csv").OrderBy(static file => file, StringComparer.Ordinal)) {
            foreach (Dictionary<string, string> row in ReadCsv(file)) {
                if (!row.TryGetValue("SHA256", out string? sha) || !row.TryGetValue("QpdfExitCode", out string? code)) continue;
                if (int.TryParse(code, out int exit)) result.TryAdd(sha.ToUpperInvariant(), exit);
            }
        }
        return result;
    }

    internal static IEnumerable<Dictionary<string, string>> ReadCsv(string path) {
        using var reader = new StreamReader(path, Encoding.UTF8);
        string? header = reader.ReadLine();
        if (header is null) yield break;
        string[] columns = SplitCsv(header);
        string? line;
        while ((line = reader.ReadLine()) != null) {
            if (line.Length == 0) continue;
            string[] values = SplitCsv(line);
            var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < columns.Length && i < values.Length; i++) row[columns[i]] = values[i];
            yield return row;
        }
    }

    private static string[] SplitCsv(string line) {
        var values = new List<string>();
        var current = new StringBuilder();
        bool quoted = false;
        for (int i = 0; i < line.Length; i++) {
            char ch = line[i];
            if (quoted) {
                if (ch == '"' && i + 1 < line.Length && line[i + 1] == '"') { current.Append('"'); i++; }
                else if (ch == '"') quoted = false;
                else current.Append(ch);
            } else if (ch == '"') {
                quoted = true;
            } else if (ch == ',') {
                values.Add(current.ToString());
                current.Clear();
            } else {
                current.Append(ch);
            }
        }
        values.Add(current.ToString());
        return values.ToArray();
    }

    internal static string Normalize(string path) => path.Replace('\\', '/');
}
