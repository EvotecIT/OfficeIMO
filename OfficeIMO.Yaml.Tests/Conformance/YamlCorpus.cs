using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>One real-world corpus file recorded in <c>Corpus/manifest.json</c>.</summary>
public sealed class YamlCorpusFile {
    internal YamlCorpusFile(string path, string category, string fullPath, byte[] bytes, IReadOnlyList<string> features, string? derivedFrom) {
        Path = path;
        Category = category;
        FullPath = fullPath;
        Bytes = bytes;
        Features = features;
        DerivedFrom = derivedFrom;
        Yaml = path.EndsWith(".md", StringComparison.OrdinalIgnoreCase) ? ExtractFrontMatter(bytes, path) : bytes;
    }

    /// <summary>Manifest path relative to the corpus folder, with forward slashes.</summary>
    public string Path { get; }

    /// <summary>Corpus category, the first path segment.</summary>
    public string Category { get; }

    /// <summary>Absolute file path.</summary>
    public string FullPath { get; }

    /// <summary>File bytes, verified against the manifest SHA-256.</summary>
    public byte[] Bytes { get; }

    /// <summary>
    /// The YAML stream in the file: the whole file, or for Markdown the front matter between the leading
    /// <c>---</c> line and the next <c>---</c> or <c>...</c> line, without those delimiter lines.
    /// </summary>
    public byte[] Yaml { get; }

    /// <summary>Feature tags from the manifest, for example <c>anchors</c> or <c>crlf</c>.</summary>
    public IReadOnlyList<string> Features { get; }

    /// <summary>Corpus path of the upstream file this one was derived from, or null for upstream bytes.</summary>
    public string? DerivedFrom { get; }

    /// <summary>Returns the corpus path.</summary>
    public override string ToString() => Path;

    private static byte[] ExtractFrontMatter(byte[] bytes, string path) {
        int firstEnd = LineEnd(bytes, 0, out int next);
        if (!IsLine(bytes, 0, firstEnd, "---")) throw new InvalidDataException(path + " does not start with a front matter line.");
        int start = next;
        int position = next;
        while (position < bytes.Length) {
            int end = LineEnd(bytes, position, out next);
            if (IsLine(bytes, position, end, "---") || IsLine(bytes, position, end, "...")) {
                var yaml = new byte[position - start];
                Array.Copy(bytes, start, yaml, 0, yaml.Length);
                return yaml;
            }

            position = next;
        }

        throw new InvalidDataException(path + " has no closing front matter line.");
    }

    // Returns the index of the line terminator (or end of data); next receives the start of the following line.
    private static int LineEnd(byte[] bytes, int start, out int next) {
        int i = start;
        while (i < bytes.Length && bytes[i] != (byte)'\n') i++;
        next = i < bytes.Length ? i + 1 : i;
        return i > start && bytes[i - 1] == (byte)'\r' ? i - 1 : i;
    }

    private static bool IsLine(byte[] bytes, int start, int end, string text) {
        if (end - start != text.Length) return false;
        for (int i = 0; i < text.Length; i++) {
            if (bytes[start + i] != (byte)text[i]) return false;
        }

        return true;
    }
}

/// <summary>The real-world corpus under <c>OfficeIMO.TestAssets/Yaml/Corpus</c>.</summary>
public static class YamlCorpus {
    /// <summary>Loads every manifest entry and verifies each file's SHA-256.</summary>
    public static IReadOnlyList<YamlCorpusFile> Load(string? directory = null) {
        directory ??= System.IO.Path.Combine(YamlAssetPaths.Root, "Corpus");
        YamlConformanceNode manifest = YamlJsonReader.Read(File.ReadAllText(System.IO.Path.Combine(directory, "manifest.json"), Encoding.UTF8));
        YamlConformanceNode entries = manifest.Get("files") ?? throw new InvalidDataException("Corpus manifest has no files array.");
        var files = new List<YamlCorpusFile>();
        foreach (YamlConformanceNode entry in entries.Items) {
            string path = entry.GetString("path");
            string fullPath = System.IO.Path.Combine(directory, path.Replace('/', System.IO.Path.DirectorySeparatorChar));
            byte[] bytes = File.ReadAllBytes(fullPath);
            string actual = Sha256(bytes);
            string expected = entry.GetString("sha256");
            if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException("Corpus file " + path + " has SHA-256 " + actual + " but the manifest records " + expected + ".");
            }

            var features = new List<string>();
            foreach (YamlConformanceNode feature in entry.Get("features")?.Items ?? new YamlConformanceNode[0]) features.Add(feature.Scalar!);
            files.Add(new YamlCorpusFile(path, entry.GetString("category"), fullPath, bytes, features, entry.Get("derivedFrom")?.Scalar));
        }

        return files;
    }

    /// <summary>Lower-case hexadecimal SHA-256 of the bytes.</summary>
    public static string Sha256(byte[] bytes) {
        using SHA256 sha = SHA256.Create();
        byte[] hash = sha.ComputeHash(bytes);
        var text = new StringBuilder(hash.Length * 2);
        foreach (byte b in hash) text.Append(b.ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
        return text.ToString();
    }
}
