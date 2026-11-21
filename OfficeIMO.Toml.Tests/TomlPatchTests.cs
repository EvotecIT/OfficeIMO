using Xunit;

namespace OfficeIMO.Toml.Tests;

// RFC 6902 JSON Patch and RFC 7386 JSON Merge Patch on the TOML model: judged by the saved text (only the patched lines
// change) and by reloading it, and a failed patch must leave the document's text as it was.
public class TomlPatchTests {
    private static (string Saved, TomlPatchReport Report) Patch(string toml, string patch) {
        using TomlDocument document = TomlDocument.Parse(toml);
        TomlPatchReport report = document.ApplyJsonPatch(patch);
        string saved = document.ToString();
        TomlDocument.Parse(saved).Dispose();
        return (saved, report);
    }

    private static (string Saved, TomlPatchReport Report) MergePatch(string toml, string patch) {
        using TomlDocument document = TomlDocument.Parse(toml);
        TomlPatchReport report = document.ApplyMergePatch(patch);
        string saved = document.ToString();
        TomlDocument.Parse(saved).Dispose();
        return (saved, report);
    }

    [Fact]
    public void APatchChangesOnlyTheLinesItTouches() {
        const string Toml = "# Project settings\n[project]\nname = 'demo'\nversion = '1.4.0' # bumped by CI\ndependencies = [\n  'httpx>=0.27',\n  'rich',\n]\n\n[tool.ruff]\nline-length = 88\nexclude = ['build']\n";

        (string saved, TomlPatchReport report) = Patch(Toml, """
            [
              { "op": "test", "path": "/project/version", "value": "1.4.0" },
              { "op": "replace", "path": "/project/version", "value": "1.5.0" },
              { "op": "add", "path": "/project/dependencies/-", "value": "click" },
              { "op": "remove", "path": "/tool/ruff/exclude" },
              { "op": "add", "path": "/tool/ruff/lint", "value": { "select": ["E", "F"] } }
            ]
            """);

        Assert.Equal("# Project settings\n[project]\nname = 'demo'\nversion = '1.5.0' # bumped by CI\ndependencies = [\n  'httpx>=0.27',\n  'rich',\n  'click',\n]\n\n[tool.ruff]\nline-length = 88\n\n[tool.ruff.lint]\nselect = [\"E\", \"F\"]\n", saved);
        Assert.Equal(new[] { "project.dependencies[2]", "tool.ruff.lint" }, report.Added);
        Assert.Equal(new[] { "project.version" }, report.Replaced);
        Assert.Equal(new[] { "tool.ruff.exclude" }, report.Removed);
    }

    [Fact]
    public void ArraysOfTablesTakeTablesAtTheirEnd() {
        const string Toml = "[[package]]\nname = \"anyio\"\nversion = \"4.3.0\"\n";

        (string saved, TomlPatchReport report) = Patch(Toml, """
            [
              { "op": "add", "path": "/package/-", "value": { "name": "idna", "version": "3.7" } },
              { "op": "replace", "path": "/package/0/version", "value": "4.4.0" }
            ]
            """);

        Assert.Equal("[[package]]\nname = \"anyio\"\nversion = \"4.4.0\"\n\n[[package]]\nname = \"idna\"\nversion = \"3.7\"\n", saved);
        Assert.Equal(new[] { "package[1]" }, report.Added);

        using TomlDocument document = TomlDocument.Parse(Toml);
        TomlPatchException exception = Assert.Throws<TomlPatchException>(() => document.ApplyJsonPatch("""[ { "op": "add", "path": "/package/0", "value": { "name": "first" } } ]"""));
        Assert.Contains("appended", exception.Message);
    }

    [Fact]
    public void MoveAndCopyCarryValuesWithTheirTypes() {
        const string Toml = "[build]\nreleased = 2024-05-01T10:00:00Z\nflags = [1, 2]\n\n[old]\nname = \"x\"\n";

        (string saved, _) = Patch(Toml, """
            [
              { "op": "copy", "from": "/build/released", "path": "/build/published" },
              { "op": "move", "from": "/old/name", "path": "/build/name" },
              { "op": "test", "path": "/build/published", "value": "2024-05-01T10:00:00Z" },
              { "op": "test", "path": "/build/flags", "value": [1.0, 2] }
            ]
            """);

        using TomlDocument reloaded = TomlDocument.Parse(saved);
        Assert.Equal(TomlNodeKind.DateTime, reloaded.Select("build.published")!.Kind);
        Assert.Equal("x", reloaded.Select("build.name")!.AsString());
        Assert.Empty(reloaded.Select("old")!.AsTable());
    }

    [Theory]
    [InlineData("""[ { "op": "replace", "path": "/version", "value": "2" }, { "op": "test", "path": "/name", "value": "other" } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/version", "value": "2" }, { "op": "add", "path": "/name", "value": null } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/version", "value": "2" }, { "op": "add", "path": "/tags/5", "value": "x" } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/version", "value": "2" }, { "op": "remove", "path": "/missing" } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/version", "value": "2" }, { "op": "add", "path": "", "value": [1] } ]""", 1)]
    [InlineData("""[ { "op": "add", "path": "version", "value": "2" } ]""", 0)]
    [InlineData("""[ { "op": "add", "path": "/a", "value": 1, "op": "remove" } ]""", -1)]
    [InlineData("""[ { "op": "add", "path": "/a", "value": 1 }""", -1)]
    public void AFailedPatchLeavesTheDocumentUnchanged(string patch, int failedOperation) {
        const string Toml = "name = \"demo\" # the name\nversion = \"1\"\ntags = [\"a\"]\n";
        using TomlDocument document = TomlDocument.Parse(Toml);

        TomlPatchException exception = Assert.Throws<TomlPatchException>(() => document.ApplyJsonPatch(patch));

        Assert.Equal(failedOperation, exception.OperationIndex);
        Assert.Equal(Toml, document.ToString());
    }

    [Fact]
    public void RemovingTheLastKeyOrTableKeepsTheContainer() {
        (string saved, _) = Patch("a.x = 1\nb = 2\n", """[ { "op": "remove", "path": "/a/x" } ]""");
        using (TomlDocument reloaded = TomlDocument.Parse(saved)) Assert.Empty(reloaded.Select("a")!.AsTable());

        (saved, _) = MergePatch("a.b.x = 1\n", """{ "a": { "b": { "x": null } } }""");
        using (TomlDocument reloaded = TomlDocument.Parse(saved)) Assert.Empty(reloaded.Select("a.b")!.AsTable());

        (saved, _) = Patch("name = \"lock\"\n\n[[package]]\nname = \"anyio\"\n", """[ { "op": "remove", "path": "/package/0" } ]""");
        using (TomlDocument reloaded = TomlDocument.Parse(saved)) Assert.Empty(reloaded.Select("package")!.AsArray());
    }

    [Theory]
    [InlineData("""[ { "op": "remove", "path": "/missing/x" } ]""")]
    [InlineData("""[ { "op": "test", "path": "/n", "value": 9007199254740992.0 } ]""")]
    [InlineData("""[ { "op": "add", "path": "/big", "value": 1e400 } ]""")]
    public void RejectsMissingParentsInexactNumbersAndOverflow(string patch) {
        const string Toml = "n = 9007199254740993\n";
        using TomlDocument document = TomlDocument.Parse(Toml);
        Assert.Throws<TomlPatchException>(() => document.ApplyJsonPatch(patch));
        Assert.Equal(Toml, document.ToString());
        document.ApplyJsonPatch("""[ { "op": "test", "path": "/n", "value": 9007199254740993 } ]""");
    }

    [Fact]
    public void AMergePatchRemovesWithNullAndMergesTables() {
        const string Toml = "# Settings\ntitle = \"app\"\n\n[server]\nhost = \"localhost\" # local only\nport = 8080\ndebug = true\n\n[cache]\nsize = 10\n";

        (string saved, TomlPatchReport report) = MergePatch(Toml, """{ "server": { "port": 9090, "debug": null }, "cache": null, "logging": { "level": "info", "file": null } }""");

        Assert.Equal("# Settings\ntitle = \"app\"\n\n[server]\nhost = \"localhost\" # local only\nport = 9090\n\n[logging]\nlevel = \"info\"\n", saved);
        Assert.Equal(new[] { "server.debug", "cache" }, report.Removed);
        Assert.Equal(new[] { "server.port" }, report.Replaced);
        Assert.Equal(new[] { "logging" }, report.Added);
        Assert.Empty(report.Diagnostics);
    }

    [Fact]
    public void AMergePatchReplacesValuesOfAnotherKindAndArrays() {
        const string Toml = "mode = \"fast\"\nports = [80, 443]\n\n[extra]\nkeep = 1\n";

        (string saved, TomlPatchReport report) = MergePatch(Toml, """{ "mode": { "name": "fast", "level": null }, "ports": [80, 8443], "extra": 5 }""");

        using TomlDocument reloaded = TomlDocument.Parse(saved);
        Assert.Equal("fast", reloaded.Select("mode.name")!.AsString());
        Assert.Null(reloaded.Select("mode.level"));
        Assert.Equal(8443, reloaded.Select("ports[1]")!.GetValue<int>());
        Assert.Equal(5, reloaded.Select("extra")!.GetValue<int>());
        Assert.Empty(report.Diagnostics);
    }

    [Theory]
    [InlineData("""["a"]""")]
    [InlineData("""{ "list": [1, null] }""")]
    [InlineData("""{ "a": 1 } x""")]
    public void AMergePatchThatTomlCannotHoldFails(string patch) {
        const string Toml = "a = 1\n";
        using TomlDocument document = TomlDocument.Parse(Toml);
        Assert.Throws<TomlPatchException>(() => document.ApplyMergePatch(patch));
        Assert.Equal(Toml, document.ToString());
    }

    [Fact]
    public void AnEmptyArrayOfTablesTakesPartInTheAtomicCheck() {
        // An array of tables with no tables has no text, but a test of its parent still sees it.
        using TomlDocument document = TomlDocument.Parse("version = 1\n");
        document.Root.AddTableArray("pending");

        Assert.Throws<TomlPatchException>(() => document.ApplyJsonPatch("""[ { "op": "replace", "path": "/version", "value": 2 }, { "op": "test", "path": "", "value": { "version": 2 } } ]"""));

        Assert.Equal(1, document.Select("version")!.GetValue<int>());
    }

    [Theory]
    [InlineData("\\uD800")]
    [InlineData("\\uDC00")]
    [InlineData("\\uD800x")]
    [InlineData("\\uD800\\u0041")]
    public void UnpairedSurrogateEscapesAreRejected(string escape) {
        using TomlDocument document = TomlDocument.Parse("a = 1\n");
        Assert.Throws<TomlPatchException>(() => document.ApplyMergePatch("{ \"s\": \"" + escape + "\" }"));
        Assert.Equal("a = 1\n", document.ToString());

        document.ApplyMergePatch("{ \"s\": \"\\uD83D\\uDE00\" }");
        Assert.Equal("\U0001F600", document.Select("s")!.AsString());
    }
}
