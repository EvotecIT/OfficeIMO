using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;

namespace OfficeIMO.Yaml.Benchmarks.LargeCorpus;

/// <summary>One indexed file of the large corpus.</summary>
internal sealed class YamlLargeCorpusFile {
    public YamlLargeCorpusFile(string source, string category, string path, string kind, string fullPath, long bytes, string sha256) {
        Source = source;
        Category = category;
        Path = path;
        Kind = kind;
        FullPath = fullPath;
        Bytes = bytes;
        Sha256 = sha256;
    }

    public string Source { get; }

    public string Category { get; }

    /// <summary>Path inside the source repository.</summary>
    public string Path { get; }

    /// <summary><c>yaml</c> or <c>front-matter</c>.</summary>
    public string Kind { get; }

    public string FullPath { get; }

    public long Bytes { get; }

    public string Sha256 { get; }

    /// <summary>Short display name: source id and path.</summary>
    public string Name => Source + ":" + Path;

    /// <summary>Reads the file, checks its recorded SHA-256 and returns the YAML stream it contains.</summary>
    public byte[] ReadYaml() {
        byte[] bytes = File.ReadAllBytes(FullPath);
        string actual = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        if (!string.Equals(actual, Sha256, StringComparison.Ordinal)) throw new InvalidDataException(Name + " no longer matches the index; re-run the fetcher.");
        return Kind == "front-matter" ? YamlLargeCorpus.ExtractFrontMatter(bytes) : bytes;
    }
}

/// <summary>A file the fetcher recorded as excluded, with its reason.</summary>
internal sealed record YamlLargeCorpusExclusion(string Source, string Category, string Path, string Reason);

/// <summary>Source summary from the index.</summary>
internal sealed record YamlLargeCorpusSource(string Id, string Category, string Repository, string Commit, string License, int Files, int Excluded);

/// <summary>
/// The opt-in large corpus fetched by <c>OfficeIMO.TestAssets/Yaml/LargeCorpus/Update-YamlLargeCorpus.ps1</c>. The cache
/// is <c>--cache</c>, else <c>OFFICEIMO_YAML_CORPUS_CACHE</c>, else <c>yaml-corpus</c> under the temporary folder.
/// </summary>
internal sealed class YamlLargeCorpus {
    private YamlLargeCorpus(string root, IReadOnlyList<YamlLargeCorpusSource> sources, IReadOnlyList<YamlLargeCorpusFile> files, IReadOnlyList<YamlLargeCorpusExclusion> exclusions) {
        Root = root;
        Sources = sources;
        Files = files;
        Exclusions = exclusions;
    }

    public string Root { get; }

    public IReadOnlyList<YamlLargeCorpusSource> Sources { get; }

    public IReadOnlyList<YamlLargeCorpusFile> Files { get; }

    public IReadOnlyList<YamlLargeCorpusExclusion> Exclusions { get; }

    public static string ResolveRoot(string? option) {
        string? root = option ?? Environment.GetEnvironmentVariable("OFFICEIMO_YAML_CORPUS_CACHE");
        return Path.GetFullPath(string.IsNullOrWhiteSpace(root) ? Path.Combine(Path.GetTempPath(), "yaml-corpus") : root);
    }

    public static YamlLargeCorpus Load(string root) {
        string indexPath = Path.Combine(root, "index.json");
        if (!File.Exists(indexPath)) {
            throw new FileNotFoundException("No large-corpus index at " + indexPath + ". Run OfficeIMO.TestAssets/Yaml/LargeCorpus/Update-YamlLargeCorpus.ps1 first.", indexPath);
        }

        using JsonDocument index = JsonDocument.Parse(File.ReadAllBytes(indexPath));
        var sources = new List<YamlLargeCorpusSource>();
        foreach (JsonElement source in index.RootElement.GetProperty("sources").EnumerateArray()) {
            sources.Add(new YamlLargeCorpusSource(
                source.GetProperty("id").GetString()!, source.GetProperty("category").GetString()!, source.GetProperty("repository").GetString()!,
                source.GetProperty("commit").GetString()!, source.GetProperty("license").GetString()!, source.GetProperty("files").GetInt32(), source.GetProperty("excluded").GetInt32()));
        }

        var files = new List<YamlLargeCorpusFile>();
        var exclusions = new List<YamlLargeCorpusExclusion>();
        foreach (JsonElement file in index.RootElement.GetProperty("files").EnumerateArray()) {
            string source = file.GetProperty("source").GetString()!;
            string category = file.GetProperty("category").GetString()!;
            string path = file.GetProperty("path").GetString()!;
            JsonElement excluded = file.GetProperty("excluded");
            if (excluded.ValueKind == JsonValueKind.String) {
                exclusions.Add(new YamlLargeCorpusExclusion(source, category, path, excluded.GetString()!));
                continue;
            }

            string fullPath = Path.Combine(root, "sources", source, path.Replace('/', Path.DirectorySeparatorChar));
            files.Add(new YamlLargeCorpusFile(source, category, path, file.GetProperty("kind").GetString()!, fullPath, file.GetProperty("bytes").GetInt64(), file.GetProperty("sha256").GetString()!));
        }

        return new YamlLargeCorpus(root, sources, files, exclusions);
    }

    /// <summary>
    /// Front matter between a leading <c>---</c> line and the next <c>---</c> or <c>...</c> line, without the delimiter lines,
    /// the rule the committed corpus uses. A file whose front matter is never closed returns everything after the first line.
    /// </summary>
    public static byte[] ExtractFrontMatter(byte[] bytes) {
        int start = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF ? 3 : 0;
        int position = NextLine(bytes, start);
        int contentStart = position;
        while (position < bytes.Length) {
            int next = NextLine(bytes, position);
            int end = next;
            if (end > position && bytes[end - 1] == (byte)'\n') end--;
            if (end > position && bytes[end - 1] == (byte)'\r') end--;
            if (end - position == 3 && ((bytes[position] == '-' && bytes[position + 1] == '-' && bytes[position + 2] == '-') || (bytes[position] == '.' && bytes[position + 1] == '.' && bytes[position + 2] == '.'))) {
                return bytes.AsSpan(contentStart, position - contentStart).ToArray();
            }

            position = next;
        }

        return bytes.AsSpan(contentStart).ToArray();
    }

    private static int NextLine(byte[] bytes, int start) {
        int i = Array.IndexOf(bytes, (byte)'\n', start);
        return i < 0 ? bytes.Length : i + 1;
    }

    public static string Format(long value) => value.ToString("N0", CultureInfo.InvariantCulture);
}
