using System.Globalization;
using System.Text;
using OfficeIMO.Tests;

namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>One yaml-test-suite case or numbered subcase.</summary>
public sealed class YamlTestSuiteCase {
    internal YamlTestSuiteCase(string id, string directory) {
        Id = id;
        Directory = directory;
        Name = ReadOptionalText("===")?.Trim() ?? string.Empty;
        Input = File.ReadAllBytes(Path.Combine(directory, "in.yaml"));
        ExpectedEvents = ReadOptionalText("test.event") ?? throw new InvalidDataException("Case " + id + " has no test.event.");
        IsError = File.Exists(Path.Combine(directory, "error"));
        Json = ReadOptionalText("in.json");
        OutYaml = ReadOptionalText("out.yaml");
        EmitYaml = ReadOptionalText("emit.yaml");
    }

    /// <summary>Case identifier, for example <c>229Q</c> or <c>3RLN/01</c> for a subcase.</summary>
    public string Id { get; }

    /// <summary>Human-readable case name from the <c>===</c> file.</summary>
    public string Name { get; }

    /// <summary>Folder that holds the case files.</summary>
    public string Directory { get; }

    /// <summary>Raw bytes of <c>in.yaml</c>.</summary>
    public byte[] Input { get; }

    /// <summary>Expected event text from <c>test.event</c>. For error cases it lists the events before the error.</summary>
    public string ExpectedEvents { get; }

    /// <summary>The input is invalid and must be rejected.</summary>
    public bool IsError { get; }

    /// <summary>Expected JSON stream from <c>in.json</c>, or null when the suite gives none.</summary>
    public string? Json { get; }

    /// <summary>Normalized YAML output from <c>out.yaml</c>, or null.</summary>
    public string? OutYaml { get; }

    /// <summary>Emitter output from <c>emit.yaml</c>, or null.</summary>
    public string? EmitYaml { get; }

    /// <summary>Returns the identifier, which keeps xUnit theory names short.</summary>
    public override string ToString() => Id;

    private string? ReadOptionalText(string name) {
        string path = Path.Combine(Directory, name);
        return File.Exists(path) ? File.ReadAllText(path, new UTF8Encoding(false)) : null;
    }
}

/// <summary>The pinned yaml-test-suite data under <c>OfficeIMO.TestAssets/Yaml/yaml-test-suite</c>.</summary>
public sealed class YamlTestSuite {
    private YamlTestSuite(string directory, string commit, IReadOnlyList<YamlTestSuiteCase> cases) {
        Directory = directory;
        Commit = commit;
        Cases = cases;
    }

    /// <summary>Suite folder.</summary>
    public string Directory { get; }

    /// <summary>Upstream commit recorded in the manifest.</summary>
    public string Commit { get; }

    /// <summary>Every case and subcase, ordered by identifier.</summary>
    public IReadOnlyList<YamlTestSuiteCase> Cases { get; }

    /// <summary>Loads the suite and checks the case counts against its manifest.</summary>
    public static YamlTestSuite Load(string? directory = null) {
        directory ??= Path.Combine(YamlAssetPaths.Root, "yaml-test-suite");
        YamlConformanceNode manifest = YamlJsonReader.Read(File.ReadAllText(Path.Combine(directory, "manifest.json")));
        var cases = new List<YamlTestSuiteCase>();
        foreach (string caseDirectory in SortedDirectories(directory)) {
            string id = Path.GetFileName(caseDirectory);
            if (!IsCaseId(id)) continue;
            if (File.Exists(Path.Combine(caseDirectory, "in.yaml"))) cases.Add(new YamlTestSuiteCase(id, caseDirectory));
            foreach (string subcase in SortedDirectories(caseDirectory)) {
                string number = Path.GetFileName(subcase);
                if (IsDigits(number) && File.Exists(Path.Combine(subcase, "in.yaml"))) cases.Add(new YamlTestSuiteCase(id + "/" + number, subcase));
            }
        }

        int expected = int.Parse(manifest.Get("caseCount")?.Scalar ?? "-1", CultureInfo.InvariantCulture);
        if (cases.Count != expected) {
            throw new InvalidDataException("yaml-test-suite manifest lists " + expected + " cases but " + cases.Count + " were found in " + directory + ".");
        }

        return new YamlTestSuite(directory, manifest.GetString("commit"), cases);
    }

    private static IEnumerable<string> SortedDirectories(string directory) {
        string[] directories = System.IO.Directory.GetDirectories(directory);
        Array.Sort(directories, StringComparer.Ordinal);
        return directories;
    }

    // Case folders are four characters of digits and upper-case letters.
    private static bool IsCaseId(string name) {
        if (name.Length != 4) return false;
        foreach (char c in name) {
            if (!(c >= '0' && c <= '9') && !(c >= 'A' && c <= 'Z')) return false;
        }

        return true;
    }

    private static bool IsDigits(string name) {
        if (name.Length == 0) return false;
        foreach (char c in name) {
            if (c < '0' || c > '9') return false;
        }

        return true;
    }
}

/// <summary>Locates the YAML test assets in the source checkout.</summary>
public static class YamlAssetPaths {
    /// <summary><c>OfficeIMO.TestAssets/Yaml</c> in the current checkout.</summary>
    public static string Root => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Yaml");
}
