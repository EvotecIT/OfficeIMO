using System.Text;
using OfficeIMO.Toml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Toml.Tests;

// Scripted edits on real files: each asserts the exact lines the save removes and adds, so any reformatting,
// lost comment or moved line elsewhere in the file fails. The oracle cannot reach this; it only reads.
public class TomlCorpusEditTests {
    [Fact]
    public void BumpingAVersionChangesOnlyThatLineAndKeepsItsComment() {
        AssertEdit("Python/sampleproject-pyproject.toml",
            doc => doc.Set("project.version", "4.0.1"),
            removed: new[] { "version = \"4.0.0\" # REQUIRED, although can be dynamic" },
            added: new[] { "version = \"4.0.1\" # REQUIRED, although can be dynamic" });
    }

    [Fact]
    public void AddingADependencyAddsOneLineAfterTheLastDependency() {
        AssertEdit("Cargo/ripgrep-Cargo.toml",
            doc => doc.Select("dependencies")!.AsTable().Add("serde", "1.0"),
            removed: Array.Empty<string>(),
            added: new[] { "serde = \"1.0\"" },
            after: "textwrap = { version = \"0.16.0\", default-features = false }");
    }

    [Fact]
    public void AddingAWorkspaceMemberFollowsTheMultiLineArrayLayout() {
        AssertEdit("Cargo/ripgrep-Cargo.toml",
            doc => doc.Select("workspace.members")!.AsArray().Add("crates/toml"),
            removed: Array.Empty<string>(),
            added: new[] { "  \"crates/toml\"," },
            after: "  \"crates/ignore\",");
    }

    [Fact]
    public void SettingAnAlignedValueKeepsTheAlignment() {
        AssertEdit("Hugo/hugo-docs-hugo.toml",
            doc => doc.Set("caches.images.maxAge", "720h"),
            removed: new[] { "    maxAge = \"1440h\"" },
            added: new[] { "    maxAge = \"720h\"" });
    }

    [Fact]
    public void AddingAKeyToAnIndentedTableUsesItsIndentation() {
        AssertEdit("Hugo/hugo-docs-hugo.toml",
            doc => doc.Select("caches.images")!.AsTable().Add("enabled", true),
            removed: Array.Empty<string>(),
            added: new[] { "    enabled = true" },
            after: "    maxAge = \"1440h\"");
    }

    [Fact]
    public void RemovingALockFilePackageDeletesOnlyItsSection() {
        TomlCorpusFile file = Corpus("Uv/uv-uv.lock");
        string original = Decode(file.Bytes);
        using TomlDocument doc = TomlDocument.Load(new MemoryStream(file.Bytes));
        TomlTableArray packages = doc.Select("package")!.AsTableArray();
        int index = packages.ToList().FindIndex(p => p["name"].AsString() == "anyio");
        packages.RemoveAt(index);
        string edited = Save(doc);

        int start = original.IndexOf("[[package]]\nname = \"anyio\"\n", StringComparison.Ordinal);
        int end = original.IndexOf("[[package]]", start + 1, StringComparison.Ordinal);
        Assert.Equal(original.Remove(start, end - start), edited);
        Assert.DoesNotContain(TomlDocument.Parse(edited).Select("package")!.AsTableArray(), p => p["name"].AsString() == "anyio");
    }

    [Fact]
    public void RemovingAStarshipModuleDeletesItsSectionAndNothingElse() {
        TomlCorpusFile file = Corpus("Starship/nerd-font-symbols.toml");
        string original = Decode(file.Bytes);
        using TomlDocument doc = TomlDocument.Load(new MemoryStream(file.Bytes));
        Assert.True(doc.Remove("azure"));
        string edited = Save(doc);
        int start = original.IndexOf("[azure]", StringComparison.Ordinal);
        int end = original.IndexOf("[battery]", StringComparison.Ordinal);
        Assert.Equal(original.Remove(start, end - start), edited);
    }

    [Fact]
    public void SettingAMissingToolTableAddsOneSectionAtTheEnd() {
        TomlCorpusFile file = Corpus("Python/ruff-pyproject.toml");
        string original = Decode(file.Bytes);
        using TomlDocument doc = TomlDocument.Load(new MemoryStream(file.Bytes));
        doc.Set("tool.officeimo.enabled", true);
        string edited = Save(doc);
        Assert.StartsWith(original, edited, StringComparison.Ordinal);
        Assert.Equal("\n[tool.officeimo]\nenabled = true\n", edited.Substring(original.Length));
    }

    private static void AssertEdit(string path, Action<TomlDocument> edit, string[] removed, string[] added, string? after = null) {
        TomlCorpusFile file = Corpus(path);
        string original = Decode(file.Bytes);
        using TomlDocument doc = TomlDocument.Load(new MemoryStream(file.Bytes));
        edit(doc);
        string edited = Save(doc);
        (List<string> actualRemoved, List<string> actualAdded, int addedAt) = LineDiff(original, edited);
        Assert.Equal(removed, actualRemoved);
        Assert.Equal(added, actualAdded);
        if (after != null) Assert.Equal(after, original.Split('\n')[addedAt - 1]);
        TomlDocument.Parse(edited);
    }

    private static TomlCorpusFile Corpus(string path) => TomlCorpus.Load().Single(f => f.Path == path);

    private static string Decode(byte[] bytes) => new UTF8Encoding(false).GetString(bytes);

    private static string Save(TomlDocument doc) {
        var output = new MemoryStream();
        doc.Save(output);
        return Decode(output.ToArray());
    }

    /// <summary>Returns the lines only in the original, the lines only in the edit, and the original line index where the change starts.</summary>
    private static (List<string> Removed, List<string> Added, int At) LineDiff(string original, string edited) {
        string[] a = original.Split('\n'), b = edited.Split('\n');
        int prefix = 0;
        while (prefix < a.Length && prefix < b.Length && a[prefix] == b[prefix]) prefix++;
        int suffix = 0;
        while (suffix < a.Length - prefix && suffix < b.Length - prefix && a[a.Length - 1 - suffix] == b[b.Length - 1 - suffix]) suffix++;
        return (a.Skip(prefix).Take(a.Length - prefix - suffix).ToList(), b.Skip(prefix).Take(b.Length - prefix - suffix).ToList(), prefix);
    }
}
