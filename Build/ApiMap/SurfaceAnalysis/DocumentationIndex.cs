using System.Text.RegularExpressions;

namespace OfficeIMO.Build.ApiMap;

internal sealed class DocumentationIndex {
    private readonly Dictionary<string, List<string>> _identifiers = new(StringComparer.Ordinal);
    private readonly Dictionary<string, List<string>> _construction = new(StringComparer.Ordinal);
    internal int FileCount { get; private set; }

    internal DocumentationIndex(string root, string[] files, string package) {
        string format = package == "OfficeIMO" ? "" : package["OfficeIMO.".Length..].Split('.')[0];
        foreach (string file in files.Where(file => IsEvidence(file, package, format))) {
            string path = Path.Combine(root, file);
            if (!File.Exists(path)) continue;
            FileCount++;
            string text = File.ReadAllText(path);
            foreach (string identifier in Regex.Matches(text, @"\b[A-Za-z_][A-Za-z0-9_]*\b").Select(match => match.Value).Distinct(StringComparer.Ordinal)) {
                if (!_identifiers.TryGetValue(identifier, out var evidence)) _identifiers[identifier] = evidence = [];
                evidence.Add(file);
            }
            foreach (Match match in Regex.Matches(text,
                @"\bnew\s+(?:global::)?(?:[A-Za-z_]\w*\.)*(?<type>[A-Za-z_]\w*)\b|\[(?:[A-Za-z_]\w*\.)*(?<type>[A-Za-z_]\w*)\]::new\b")) {
                string name = match.Groups["type"].Value;
                if (!_construction.TryGetValue(name, out var evidence)) _construction[name] = evidence = [];
                if (!evidence.Contains(file)) evidence.Add(file);
            }
        }
    }

    internal IReadOnlyList<string> Find(string identifier) =>
        _identifiers.TryGetValue(identifier, out var evidence) ? evidence : Array.Empty<string>();

    internal IReadOnlyList<string> FindConstruction(string identifier) =>
        _construction.TryGetValue(identifier, out var evidence) ? evidence : Array.Empty<string>();

    private static bool IsEvidence(string file, string package, string format) {
        if (file.EndsWith(".cs", StringComparison.OrdinalIgnoreCase))
            return file.StartsWith("OfficeIMO.Examples/", StringComparison.Ordinal)
                || file.StartsWith("OfficeIMO.MarkdownRenderer.SamplePlugin/", StringComparison.Ordinal);
        if (file.EndsWith(".ps1", StringComparison.OrdinalIgnoreCase))
            return file.StartsWith("PowerShell/PSWriteOffice/Examples/", StringComparison.Ordinal);
        if (!file.EndsWith(".md", StringComparison.OrdinalIgnoreCase)) return false;
        if (file.StartsWith("OfficeIMO", StringComparison.Ordinal) && file.EndsWith("/README.md", StringComparison.Ordinal)) return true;
        if (file.Equals("README.md", StringComparison.Ordinal) || file.Equals($"{package}/README.md", StringComparison.Ordinal)) return true;
        if (file.StartsWith($"OfficeIMO.{format}.", StringComparison.Ordinal) && file.EndsWith("/README.md", StringComparison.Ordinal)) return true;
        if (!file.StartsWith("Docs/", StringComparison.Ordinal) && !file.StartsWith("Website/content/docs/", StringComparison.Ordinal)) return false;
        if (file.Contains("/generated/", StringComparison.OrdinalIgnoreCase) || file.Contains("/api/", StringComparison.OrdinalIgnoreCase) || file.Contains("/reference/", StringComparison.OrdinalIgnoreCase)) return false;
        string name = Path.GetFileName(file);
        if (name.Equals("ROADMAP.md", StringComparison.OrdinalIgnoreCase) || name.Contains("design", StringComparison.OrdinalIgnoreCase)
            || name.Contains("migration", StringComparison.OrdinalIgnoreCase) || name.Contains("benchmark", StringComparison.OrdinalIgnoreCase)) return false;
        return package == "OfficeIMO.Core" || package == "OfficeIMO" || file.Contains(format, StringComparison.OrdinalIgnoreCase);
    }
}
