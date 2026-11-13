using Xunit;

namespace OfficeIMO.Toml.Tests;

// Edit rules the oracle cannot check: what text a save writes for each kind of change.
public class TomlEditingTests {
    private static string Edit(string toml, Action<TomlDocument> edit) {
        using TomlDocument doc = TomlDocument.Parse(toml);
        edit(doc);
        string saved = doc.ToString();
        TomlDocument.Parse(saved);
        return saved;
    }

    [Theory]
    [InlineData("n = 1_000\n", 1000L)]
    [InlineData("n = 0xFF\n", 255L)]
    [InlineData("f = 1e3\n", 1000.0)]
    [InlineData("s = 'x'\n", "x")]
    [InlineData("d = 1979-05-27 07:32:00Z\n", null)]
    public void AssigningAnEqualValueKeepsTheSpelling(string toml, object? value) {
        string key = toml.Substring(0, 1);
        Assert.Equal(toml, Edit(toml, doc => {
            TomlValue node = doc.Root[key].AsValue();
            node.SetValue(value ?? node.Value);
        }));
    }

    [Theory]
    [InlineData("n = 0xff # mask\n", 4096L, "n = 0x1000 # mask\n")]
    [InlineData("n = 0o755\n", 8L, "n = 0o10\n")]
    [InlineData("n = 0xFF\n", -1L, "n = -1\n")]
    [InlineData("s = 'path'\n", "other", "s = 'other'\n")]
    [InlineData("s = 'path'\n", "it's", "s = \"it's\"\n")]
    [InlineData("s = \"\"\"\na\"\"\"\n", "b\nc", "s = \"\"\"\nb\nc\"\"\"\n")]
    [InlineData("d = 1979-05-27 07:32:00Z\n", "2001-02-03", "d = 2001-02-03 07:32:00Z\n")]
    [InlineData("f = 1.5\n", double.NaN, "f = nan\n")]
    public void AChangedValueKeepsItsStyleWhenTheStyleCanHoldIt(string toml, object value, string expected) {
        if (value is "2001-02-03") value = new DateTimeOffset(2001, 2, 3, 7, 32, 0, TimeSpan.Zero);
        Assert.Equal(expected, Edit(toml, doc => doc.Root[doc.Root.Keys.First()].AsValue().SetValue(value)));
    }

    [Fact]
    public void NewStringsFollowTheQuotingOfTheirSiblings() {
        Assert.Equal("a = 'x'\nb = 'y'\nc = 'z'\n", Edit("a = 'x'\nb = 'y'\n", doc => doc.Root.Add("c", "z")));
        Assert.Equal("a = [\"x\", \"y\"]\n", Edit("a = [\"x\"]\n", doc => doc.Root["a"].AsArray().Add("y")));
    }

    [Fact]
    public void NewLinesUseTheDocumentLineEndings() {
        Assert.Equal("a = 1\r\nb = 2\r\n[t]\r\nx = 1\r\ny = 2\r\n",
            Edit("a = 1\r\n[t]\r\nx = 1\r\n", doc => {
                doc.Root.Add("b", 2);
                doc.Root["t"].AsTable().Add("y", 2);
            }));
    }

    [Fact]
    public void RemovingAKeyRemovesTheCommentLinesDirectlyAboveIt() {
        Assert.Equal("# keep\n\na = 1\n", Edit("# keep\n\na = 1\n# about b\nb = 2 # trailing\n", doc => doc.Root.Remove("b")));
    }

    // A table that has text only through its dotted keys or deeper headers disappears with its last key, in the model
    // as in the saved text (found by the large corpus: `repository.workspace = true` in Cargo manifests).
    [Theory]
    [InlineData("[package]\nname = \"a\"\nrepository.workspace = true\n", "package.repository.workspace", "package.repository", "[package]\nname = \"a\"\n")]
    [InlineData("[a.b]\nx = 1\n", "a.b", "a", "")]
    public void RemovingTheLastKeyOfATableWithoutTextRemovesTheTable(string toml, string path, string emptied, string expected) {
        Assert.Equal(expected, Edit(toml, doc => {
            Assert.True(doc.Remove(path));
            Assert.Null(doc.Select(emptied));
        }));
    }

    // Removing the last line of a file without a final line break must not leave an empty line for added text to follow.
    [Theory]
    [InlineData("a = 1\nb = 2", false, "a = 1")]
    [InlineData("a = 1\r\nb = 2", false, "a = 1")]
    [InlineData("a = 1\nb = 2", true, "a = 1\n\n[t]\nk = 1\n")]
    public void RemovingTheLastLineOfAFileWithoutAFinalBreakTakesTheBreakBeforeIt(string toml, bool addTable, string expected) {
        Assert.Equal(expected, Edit(toml, doc => {
            doc.Remove("b");
            if (addTable) doc.Root.AddTable("t").Add("k", 1);
        }));
    }

    [Fact]
    public void RemovingATableRemovesItsSubTablesWhereverTheyAre() {
        string toml = "[a]\nx = 1\n\n[b]\ny = 2\n\n[a.c]\nz = 3\n";
        Assert.Equal("[b]\ny = 2\n", Edit(toml, doc => doc.Remove("a")));
    }

    [Fact]
    public void AddingToAnImplicitTableGivesItAHeaderBeforeItsSubTables() {
        Assert.Equal("[a]\nx = 1\n\n[a.b]\ny = 2\n", Edit("[a.b]\ny = 2\n", doc => doc.Root["a"].AsTable().Add("x", 1)));
    }

    [Fact]
    public void AddingToADottedTableWritesADottedKeyNextToItsSiblings() {
        Assert.Equal("site.name = 'x'\nsite.url = 'u'\nother = 1\n", Edit("site.name = 'x'\nother = 1\n", doc => doc.Root["site"].AsTable().Add("url", "u")));
    }

    [Theory]
    [InlineData("a = [1, 2]\n", "a = [1, 2, 3]\n")]
    [InlineData("a = [1, 2,]\n", "a = [1, 2, 3,]\n")]
    [InlineData("a = []\n", "a = [3]\n")]
    [InlineData("a = [\n  1,\n  2,\n]\n", "a = [\n  1,\n  2,\n  3,\n]\n")]
    [InlineData("a = [\n  1,\n  2  # two\n]\n", "a = [\n  1,\n  2,  # two\n  3\n]\n")]
    public void AppendedArrayItemsFollowTheArrayLayout(string toml, string expected) {
        Assert.Equal(expected, Edit(toml, doc => doc.Root["a"].AsArray().Add(3)));
    }

    [Theory]
    [InlineData("a = [1, 2, 3]\n", 0, "a = [2, 3]\n")]
    [InlineData("a = [1, 2, 3]\n", 2, "a = [1, 2]\n")]
    [InlineData("a = [\n  1, # one\n  2, # two\n]\n", 1, "a = [\n  1, # one\n]\n")]
    [InlineData("a = [\n  1, # one\n  2, # two\n]\n", 0, "a = [\n  2, # two\n]\n")]
    public void RemovedArrayItemsTakeOneSeparatorWithThem(string toml, int index, string expected) {
        Assert.Equal(expected, Edit(toml, doc => doc.Root["a"].AsArray().RemoveAt(index)));
    }

    [Fact]
    public void InlineTablesAreEditedInsideTheirBraces() {
        Assert.Equal("p = { x = 1, y = 2, z = 3 }\n", Edit("p = { x = 1, y = 2 }\n", doc => doc.Root["p"].AsTable().Add("z", 3)));
        Assert.Equal("p = { y = 2 }\n", Edit("p = { x = 1, y = 2 }\n", doc => doc.Root["p"].AsTable().Remove("x")));
        Assert.Equal("p = { k = \"v\" }\n", Edit("p = {}\n", doc => doc.Root["p"].AsTable().Add("k", "v")));
        Assert.Equal("p = { a.x = 1, a.y = 2 }\n", Edit("p = { a.x = 1 }\n", doc => doc.Root["p"].AsTable()["a"].AsTable().Add("y", 2)));
    }

    [Fact]
    public void NewTablesAndArraysOfTablesAreWrittenAsSections() {
        Assert.Equal("[[bin]]\nname = \"a\"\n\n[[bin]]\nname = \"b\"\n",
            Edit("[[bin]]\nname = \"a\"\n", doc => doc.Root["bin"].AsTableArray().AddTable().Add("name", "b")));
        Assert.Equal("[tool.a]\nx = 1\n\n[tool.b]\ny = 2\n\n[other]\n",
            Edit("[tool.a]\nx = 1\n\n[other]\n", doc => doc.Root["tool"].AsTable().AddTable("b").Add("y", 2)));
    }

    [Fact]
    public void ReplacingAValueWithAnotherKindWritesItInPlace() {
        Assert.Equal("v = [\"a\", \"b\"] # c\nw = 2\n", Edit("v = \"x\" # c\nw = 2\n", doc => doc.Root.Set("v", new[] { "a", "b" })));
    }

    [Fact]
    public void CommentsCanBeSetChangedAndRemoved() {
        Assert.Equal("a = 1  # one\nb = 2  # two\n", Edit("a = 1  # one\nb = 2\n", doc => doc.Root["b"].SetComment("two")));
        Assert.Equal("a = 1  # uno\n", Edit("a = 1  # one\n", doc => doc.Root["a"].SetComment("uno")));
        Assert.Equal("a = 1\n", Edit("a = 1  # one\n", doc => doc.Root["a"].SetComment(null)));
        Assert.Equal("[t] # table\n", Edit("[t]\n", doc => doc.Root["t"].SetComment("table")));
        Assert.Throws<InvalidOperationException>(() => TomlDocument.Parse("a = [1]\n").Root["a"].AsArray()[0].SetComment("x"));
    }

    [Fact]
    public void ANewDocumentIsWrittenWithBlankLinesBetweenSections() {
        using TomlDocument doc = TomlDocument.Create();
        doc.Root.Add("title", "Example");
        doc.Root.AddTable("owner").Add("dob", new DateTimeOffset(1979, 5, 27, 7, 32, 0, TimeSpan.FromHours(-8)));
        TomlTableArray products = doc.Root.AddTableArray("products");
        products.AddTable().Add("name", "Hammer").SetComment("best seller");
        products.AddTable().Add("tags", new[] { "a", "b" });
        Assert.Equal(
            "title = \"Example\"\n\n[owner]\ndob = 1979-05-27T07:32:00-08:00\n\n[[products]]\nname = \"Hammer\" # best seller\n\n[[products]]\ntags = [\"a\", \"b\"]\n",
            doc.ToString());
    }

    [Fact]
    public void PathsSelectSetAndRemove() {
        using TomlDocument doc = TomlDocument.Parse("[a]\nlist = [1, 2]\n\"dotted.key\" = 3\n[[p]]\nn = 1\n[[p]]\nn = 2\n");
        Assert.Equal(2L, doc.Select("p[1].n")!.GetValue<long>());
        Assert.Equal(3, doc.Select("a[\"dotted.key\"]")!.GetValue<int>());
        Assert.Null(doc.Select("a.missing.deeper"));
        doc.Set("a.list[0]", 10);
        doc.Set("b.c.d", "new");
        Assert.True(doc.Remove("p[0]"));
        Assert.False(doc.Remove("a.missing"));
        Assert.Equal("[a]\nlist = [10, 2]\n\"dotted.key\" = 3\n[[p]]\nn = 2\n[b.c]\nd = \"new\"\n", doc.ToString());
        Assert.Throws<ArgumentException>(() => doc.Select("a..b"));
    }
}
