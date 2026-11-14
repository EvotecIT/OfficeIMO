using System.Globalization;
using System.Text;
using OfficeIMO.Toml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Toml.Tests;

// Merges are judged by the saved text (only the overlay's changes appear, the base's comments and layout stay) and by
// reloading it: the saved text must hold exactly the data the merged model holds.
public class TomlMergeTests {
    [Fact]
    public void LayeringAPyprojectChangesOnlyWhatTheOverlaySets() {
        const string Base = "# Project settings\n[project]\nname = \"demo\"\nversion = \"1.0.0\" # bumped by CI\ndependencies = [\n  \"httpx>=0.27\",\n  \"rich\",\n]\n\n[tool.ruff]\nline-length = 88\n";
        const string Overlay = "[project]\nversion = \"1.1.0\"\ndependencies = [\"httpx>=0.27\", \"rich>=13\"]\n\n[tool.ruff]\nline-length = 100\n\n[tool.ruff.lint]\nselect = [\"E\", \"F\"] # errors only\n";
        (string saved, TomlMergeReport report) = Merge(Base, Overlay);

        Assert.Equal("# Project settings\n[project]\nname = \"demo\"\nversion = \"1.1.0\" # bumped by CI\ndependencies = [\n  \"httpx>=0.27\",\n  \"rich>=13\",\n]\n\n[tool.ruff]\nline-length = 100\n\n[tool.ruff.lint]\nselect = [\"E\", \"F\"] # errors only\n", saved);
        Assert.Equal(new[] { "tool.ruff.lint" }, report.Added);
        Assert.Equal(new[] { "project.version", "project.dependencies", "tool.ruff.line-length" }, report.Replaced);
        Assert.Empty(report.Diagnostics);
    }

    [Fact]
    public void MergeByKeyUpdatesMatchingLockFilePackagesAndAppendsNewOnes() {
        const string Base = "version = 1\n\n[[package]]\nname = \"anyio\"\nversion = \"4.3.0\"\n\n[[package]]\nname = \"idna\"\nversion = \"3.6\"\n";
        const string Overlay = "[[package]]\nname = \"idna\"\nversion = \"3.7\"\n\n[[package]]\nname = \"sniffio\"\nversion = \"1.3.1\"\n";
        (string saved, TomlMergeReport report) = Merge(Base, Overlay, TomlArrayMerge.MergeByKey("name"));

        Assert.Equal("version = 1\n\n[[package]]\nname = \"anyio\"\nversion = \"4.3.0\"\n\n[[package]]\nname = \"idna\"\nversion = \"3.7\"\n\n[[package]]\nname = \"sniffio\"\nversion = \"1.3.1\"\n", saved);
        Assert.Equal(new[] { "package[2]" }, report.Added);
        Assert.Equal(new[] { "package[1].version" }, report.Replaced);
    }

    [Fact]
    public void MergeByKeyReplacesArraysWhoseItemsHaveNoKey() {
        const string Base = "[[package]]\nname = \"a\"\nfeatures = [\"std\"]\n";
        const string Overlay = "[[package]]\nname = \"a\"\nfeatures = [\"std\", \"serde\"]\n";
        (string saved, _) = Merge(Base, Overlay, TomlArrayMerge.MergeByKey("name"));
        Assert.Equal("[[package]]\nname = \"a\"\nfeatures = [\"std\", \"serde\"]\n", saved);
    }

    [Fact]
    public void AppendAddsOverlayItemsAfterTheBaseItems() {
        const string Base = "plugins = [\"a\"]\n\n[[server]]\nhost = \"one\"\n";
        const string Overlay = "plugins = [\"b\"]\n\n[[server]]\nhost = \"two\"\n";
        (string saved, TomlMergeReport report) = Merge(Base, Overlay, TomlArrayMerge.Append);
        Assert.Equal("plugins = [\"a\", \"b\"]\n\n[[server]]\nhost = \"one\"\n\n[[server]]\nhost = \"two\"\n", saved);
        Assert.Equal(new[] { "plugins[1]", "server[1]" }, report.Added);
    }

    [Fact]
    public void AKindConflictKeepsTheOverlayValueAndIsReported() {
        const string Base = "title = \"x\"\nlog = \"info\" # level\n\n[server]\nport = 80\n";
        const string Overlay = "[log]\nlevel = \"debug\"\n\n[server]\nport = [80, 443]\n";
        (string saved, TomlMergeReport report) = Merge(Base, Overlay);
        Assert.Equal("title = \"x\"\n\n[server]\nport = [80, 443]\n\n[log]\nlevel = \"debug\"\n", saved);
        Assert.Equal(new[] { "log", "server.port" }, report.Replaced);
        Assert.Equal(new[] { OfficeDiagnosticCodes.TomlMergeKindConflict, OfficeDiagnosticCodes.TomlMergeKindConflict }, report.Diagnostics.Select(d => d.Code));
        Assert.Equal("2:7", report.Diagnostics[0].Location);
    }

    [Fact]
    public void OverlayTablesMergeIntoInlineTablesAsInlineSyntax() {
        const string Base = "[dependencies]\nserde = { version = \"1\" }\n";
        const string Overlay = "[dependencies.serde]\nfeatures = [\"derive\"]\n\n[[dependencies.serde.patch]]\nurl = \"x\"\n";
        (string saved, TomlMergeReport report) = Merge(Base, Overlay);
        Assert.Equal("[dependencies]\nserde = { version = \"1\", features = [\"derive\"], patch = [{ url = \"x\" }] }\n", saved);
        Assert.Equal(new[] { "dependencies.serde.features", "dependencies.serde.patch" }, report.Added);
    }

    [Fact]
    public void DottedKeyTablesStayDottedKeys() {
        const string Base = "[project]\nname = \"demo\"\nurls.home = \"https://example.com\"\n\n[tool.other]\nx = 1\n";
        const string Overlay = "[project]\nurls.docs = \"https://docs.example.com\"\nscripts.demo = \"demo:main\"\n";
        (string saved, _) = Merge(Base, Overlay);
        Assert.Equal("[project]\nname = \"demo\"\nurls.home = \"https://example.com\"\nurls.docs = \"https://docs.example.com\"\nscripts.demo = \"demo:main\"\n\n[tool.other]\nx = 1\n", saved);
    }

    [Theory]
    [InlineData("a.b = 1\n[x]\ny = 1\n", "[a.c]\nd = 2\n")]
    [InlineData("a.b = 1\n", "[[a.list]]\nn = 1\n")]
    [InlineData("[t]\nk = { x = 1 }\n", "[t.k.deep]\nv = 1\n[[t.k.items]]\nn = 1\n")]
    [InlineData("[[p]]\nn = 1\n[p.sub]\nv = 1\n[[p]]\nn = 2\n", "[[p]]\nn = 1\n[p.sub]\nw = 2\n[[p]]\nn = 3\n[p.sub]\nv = 9\n")]
    [InlineData("[[p]]\nn = 1\n[[p]]\nn = 2\n", "p = []\n")]
    [InlineData("[[p]]\nn = 1\n", "p = [1, 2]\n")]
    [InlineData("p = [{ n = 1 }]\n", "[[p]]\nn = 2\nm.x = 3\n")]
    [InlineData("[a]\nb.c = 1\nb.d = 2\n[[q]]\nr = 1\n", "[a]\nb = {}\n[[q]]\nr = 1\ns.t = 2\n")]
    [InlineData("v = 1\n[a]\n", "v = { x = 1 }\na = 3\n[z]\nw.x.y = 1\n[z.w.q]\nr = 1\n")]
    [InlineData("x = [[1, 2], [3]]\n", "x = [[1, 9], { a = 1 }, \"s\"]\n")]
    public void TheSavedTextReloadsToTheMergedData(string baseText, string overlayText) {
        foreach (TomlArrayMerge mode in new[] { TomlArrayMerge.Replace, TomlArrayMerge.Append, TomlArrayMerge.MergeByKey("n") }) {
            using TomlDocument doc = TomlDocument.Parse(baseText);
            using TomlDocument overlay = TomlDocument.Parse(overlayText);
            doc.Merge(overlay, new TomlMergeOptions { Arrays = mode });
            string saved = doc.ToString();
            using TomlDocument reloaded = TomlDocument.Parse(saved);
            Assert.Equal(Data(doc.Root), Data(reloaded.Root));
        }
    }

    [Fact]
    public void MergingAnyCorpusFileOntoItselfChangesNothing() {
        foreach (TomlCorpusFile file in TomlCorpus.Load()) {
            using TomlDocument doc = TomlDocument.Load(new MemoryStream(file.Bytes));
            using TomlDocument overlay = TomlDocument.Load(new MemoryStream(file.Bytes));
            TomlMergeReport report = doc.Merge(overlay, new TomlMergeOptions { Arrays = TomlArrayMerge.MergeByKey("name") });
            var output = new MemoryStream();
            doc.Save(output);
            Assert.True(file.Bytes.SequenceEqual(output.ToArray()), file.Path);
            Assert.Empty(report.Added);
            Assert.Empty(report.Replaced);
        }
    }

    private static (string Saved, TomlMergeReport Report) Merge(string baseText, string overlayText, TomlArrayMerge? arrays = null) {
        using TomlDocument doc = TomlDocument.Parse(baseText);
        using TomlDocument overlay = TomlDocument.Parse(overlayText);
        TomlMergeReport report = doc.Merge(overlay, new TomlMergeOptions { Arrays = arrays ?? TomlArrayMerge.Replace });
        string saved = doc.ToString();
        using TomlDocument reloaded = TomlDocument.Parse(saved);
        Assert.Equal(Data(doc.Root), Data(reloaded.Root));
        return (saved, report);
    }

    /// <summary>The data a node holds, independent of syntax: arrays of tables and array values read alike.</summary>
    private static string Data(TomlNode node) {
        switch (node) {
            case TomlValue value:
                return value.Kind + ":" + Convert.ToString(value.Value, CultureInfo.InvariantCulture);
            case TomlTable table:
                return "{" + string.Join(",", table.OrderBy(e => e.Key, StringComparer.Ordinal).Select(e => e.Key + "=" + Data(e.Value))) + "}";
            case TomlArray array:
                return "[" + string.Join(",", array.Select(Data)) + "]";
            default:
                return "[" + string.Join(",", ((TomlTableArray)node).Select(Data)) + "]";
        }
    }
}
