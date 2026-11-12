using System.Text;

namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>One base and overlay pair from <c>Merge/manifest.json</c>.</summary>
public sealed class YamlMergePair {
    internal YamlMergePair(string id, string basePath, byte[] baseBytes, string overlayPath, byte[] overlayBytes, string sequences, bool nullRemovesKey) {
        Id = id;
        BasePath = basePath;
        Base = baseBytes;
        OverlayPath = overlayPath;
        Overlay = overlayBytes;
        Sequences = sequences;
        NullRemovesKey = nullRemovesKey;
    }

    /// <summary>Pair identifier.</summary>
    public string Id { get; }

    /// <summary>Base path relative to the merge folder.</summary>
    public string BasePath { get; }

    /// <summary>Base bytes.</summary>
    public byte[] Base { get; }

    /// <summary>Overlay path relative to the merge folder.</summary>
    public string OverlayPath { get; }

    /// <summary>Overlay bytes.</summary>
    public byte[] Overlay { get; }

    /// <summary><c>replace</c>, <c>append</c> or <c>mergeByKey:name</c>.</summary>
    public string Sequences { get; }

    /// <summary>Whether an overlay null removes the base key.</summary>
    public bool NullRemovesKey { get; }

    /// <summary>Returns the identifier.</summary>
    public override string ToString() => Id;
}

/// <summary>The merge pairs under <c>OfficeIMO.TestAssets/Yaml/Merge</c>.</summary>
public static class YamlMergeCorpus {
    /// <summary>Loads every pair, verifying the SHA-256 of each file the manifest lists.</summary>
    public static IReadOnlyList<YamlMergePair> Load() {
        string directory = Path.Combine(YamlAssetPaths.Root, "Merge");
        YamlConformanceNode manifest = YamlJsonReader.Read(File.ReadAllText(Path.Combine(directory, "manifest.json"), Encoding.UTF8));
        var hashes = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (YamlConformanceNode file in (manifest.Get("files") ?? throw new InvalidDataException("The merge manifest has no files.")).Items) {
            hashes.Add(file.GetString("path"), file.GetString("sha256"));
        }

        // Bases from the committed corpus are verified against its own manifest.
        var corpus = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (YamlCorpusFile file in YamlCorpus.Load()) corpus.Add("../Corpus/" + file.Path, file.Bytes);

        var pairs = new List<YamlMergePair>();
        foreach (YamlConformanceNode pair in (manifest.Get("pairs") ?? throw new InvalidDataException("The merge manifest has no pairs.")).Items) {
            string basePath = pair.GetString("base"), overlayPath = pair.GetString("overlay");
            pairs.Add(new YamlMergePair(pair.GetString("id"), basePath, Read(directory, basePath, hashes, corpus), overlayPath, Read(directory, overlayPath, hashes, corpus),
                pair.GetString("sequences"), pair.Get("nullRemovesKey")?.Kind == YamlConformanceNodeKind.Boolean && pair.Get("nullRemovesKey")!.Scalar == "true"));
        }

        return pairs;
    }

    private static byte[] Read(string directory, string path, Dictionary<string, string> hashes, Dictionary<string, byte[]> corpus) {
        if (corpus.TryGetValue(path, out byte[]? fromCorpus)) return fromCorpus;
        if (!hashes.TryGetValue(path, out string? expected)) throw new InvalidDataException("The merge manifest lists no hash for " + path + ".");
        byte[] bytes = File.ReadAllBytes(Path.Combine(directory, path.Replace('/', Path.DirectorySeparatorChar)));
        string actual = YamlCorpus.Sha256(bytes);
        if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Merge file " + path + " has SHA-256 " + actual + " but the manifest records " + expected + ".");
        return bytes;
    }
}
