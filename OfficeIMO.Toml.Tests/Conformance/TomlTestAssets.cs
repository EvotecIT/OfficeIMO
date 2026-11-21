using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Tests;

namespace OfficeIMO.Toml.Tests.Conformance;

/// <summary>One toml-test case: the TOML bytes and, for valid cases, the expected tagged JSON.</summary>
public sealed class TomlTestCase {
    public TomlTestCase(string name, byte[] toml, string? json) {
        Name = name;
        Toml = toml;
        Json = json;
    }

    /// <summary>The case path without extension, such as <c>valid/array/array</c>.</summary>
    public string Name { get; }

    public byte[] Toml { get; }

    public string? Json { get; }

    public override string ToString() => Name;
}

/// <summary>The pinned toml-test cases under <c>OfficeIMO.TestAssets/Toml/toml-test</c>.</summary>
public static class TomlTestSuite {
    public static string Root => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Toml");

    private static readonly Lazy<(IReadOnlyList<TomlTestCase> Valid, IReadOnlyList<TomlTestCase> Invalid)> Cases = new(LoadCases);

    public static IReadOnlyList<TomlTestCase> Valid => Cases.Value.Valid;

    public static IReadOnlyList<TomlTestCase> Invalid => Cases.Value.Invalid;

    /// <summary>Loads the case list and fails when it drifts from the counts the manifest records.</summary>
    private static (IReadOnlyList<TomlTestCase>, IReadOnlyList<TomlTestCase>) LoadCases() {
        string directory = Path.Combine(Root, "toml-test");
        JsonNode manifest = JsonText.Parse(File.ReadAllText(Path.Combine(directory, "manifest.json"), Encoding.UTF8));
        var valid = new List<TomlTestCase>();
        var invalid = new List<TomlTestCase>();
        foreach (string line in File.ReadAllLines(Path.Combine(directory, manifest.GetString("caseList")))) {
            if (line.Length == 0 || !line.EndsWith(".toml", StringComparison.Ordinal)) continue;
            string name = line.Substring(0, line.Length - 5);
            byte[] toml = File.ReadAllBytes(Path.Combine(directory, line));
            if (line.StartsWith("valid/", StringComparison.Ordinal)) {
                valid.Add(new TomlTestCase(name, toml, File.ReadAllText(Path.Combine(directory, name + ".json"), Encoding.UTF8)));
            } else {
                invalid.Add(new TomlTestCase(name, toml, null));
            }
        }
        int expectedValid = int.Parse(manifest.Get("validCases")!.Literal!);
        int expectedInvalid = int.Parse(manifest.Get("invalidCases")!.Literal!);
        if (valid.Count != expectedValid || invalid.Count != expectedInvalid) {
            throw new InvalidDataException("toml-test has " + valid.Count + " valid and " + invalid.Count + " invalid cases; the manifest records " + expectedValid + " and " + expectedInvalid + ".");
        }
        return (valid, invalid);
    }

    /// <summary>Reads a checked-in list of case names, ignoring comments and blank lines.</summary>
    public static HashSet<string> ReadList(string fileName) {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Toml.Tests", "Conformance", fileName);
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (string line in File.ReadAllLines(path)) {
            string trimmed = line.Trim();
            if (trimmed.Length == 0 || trimmed[0] == '#') continue;
            int space = trimmed.IndexOf(' ');
            names.Add(space < 0 ? trimmed : trimmed.Substring(0, space));
        }
        return names;
    }
}

/// <summary>One real-world corpus file recorded in <c>Corpus/manifest.json</c>.</summary>
public sealed class TomlCorpusFile {
    public TomlCorpusFile(string path, byte[] bytes) {
        Path = path;
        Bytes = bytes;
    }

    public string Path { get; }

    public byte[] Bytes { get; }

    public override string ToString() => Path;
}

/// <summary>The real-world corpus under <c>OfficeIMO.TestAssets/Toml/Corpus</c>.</summary>
public static class TomlCorpus {
    public static IReadOnlyList<TomlCorpusFile> Load() {
        string directory = System.IO.Path.Combine(TomlTestSuite.Root, "Corpus");
        JsonNode manifest = JsonText.Parse(File.ReadAllText(System.IO.Path.Combine(directory, "manifest.json"), Encoding.UTF8));
        var files = new List<TomlCorpusFile>();
        foreach (JsonNode entry in manifest.Get("files")!.Array!) {
            string path = entry.GetString("path");
            byte[] bytes = File.ReadAllBytes(System.IO.Path.Combine(directory, path.Replace('/', System.IO.Path.DirectorySeparatorChar)));
            string actual = Sha256(bytes);
            if (!string.Equals(actual, entry.GetString("sha256"), StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException("Corpus file " + path + " has SHA-256 " + actual + " but the manifest records " + entry.GetString("sha256") + ".");
            }
            files.Add(new TomlCorpusFile(path, bytes));
        }
        return files;
    }

    public static string Sha256(byte[] bytes) {
        using SHA256 sha = SHA256.Create();
        var text = new StringBuilder();
        foreach (byte b in sha.ComputeHash(bytes)) text.Append(b.ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
        return text.ToString();
    }
}
