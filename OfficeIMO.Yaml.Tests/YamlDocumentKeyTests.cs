using Xunit;

namespace OfficeIMO.Yaml.Tests;

// Duplicate keys, merge keys and per-document schemas are load semantics the suite's JSON cannot express.
public class YamlDocumentKeyTests {
    // Small mappings compare keys pairwise and larger ones through a dictionary; both paths must agree.
    private static string Mapping(int padding, string duplicates) {
        var text = new System.Text.StringBuilder();
        for (int i = 0; i < padding; i++) text.Append("pad").Append(i).Append(": x\n");
        return text.Append(duplicates).ToString();
    }

    [Theory]
    [InlineData(0)]
    [InlineData(20)]
    public void DuplicateKeysFailByDefault(int padding) {
        YamlException error = Assert.Throws<YamlException>(() => YamlDocument.Parse(Mapping(padding, "a: 1\n\"a\": 2\n")));
        Assert.Equal(YamlErrorCodes.DuplicateKey, error.Code);
        Assert.Equal(padding + 2, error.Line);
    }

    [Theory]
    [InlineData(0, YamlDuplicateKeyPolicy.FirstWins)]
    [InlineData(20, YamlDuplicateKeyPolicy.FirstWins)]
    [InlineData(0, YamlDuplicateKeyPolicy.LastWins)]
    [InlineData(20, YamlDuplicateKeyPolicy.LastWins)]
    public void LenientPoliciesKeepEveryEntryAndReportEachRepeat(int padding, YamlDuplicateKeyPolicy policy) {
        string yaml = Mapping(padding, "a: 1\nb: 2\na: 3\n'a': 4\n");
        YamlDocument document = YamlDocument.Parse(yaml, new YamlLoadOptions { DuplicateKeys = policy });
        YamlMapping root = document.Root!.AsMapping();
        Assert.Equal(padding + 4, root.Count);
        Assert.Equal(policy == YamlDuplicateKeyPolicy.FirstWins ? "1" : "4", root["a"].AsString());
        Assert.Equal(2, document.Diagnostics.Count);
        Assert.All(document.Diagnostics, d => Assert.Equal(YamlErrorCodes.DuplicateKey, d.Code));
        Assert.Equal((padding + 3) + ":1", document.Diagnostics[0].Location);
        Assert.Equal(yaml, document.ToString());
    }

    // Keys are equal when their values are: same resolved type and canonical value.
    [Theory]
    [InlineData("1: a\n0x1: b\n", true)]
    [InlineData("~: a\nnull: b\n", true)]
    [InlineData("true: a\nTrue: b\n", true)]
    [InlineData("[a, b]: 1\n[a, b]: 2\n", true)]
    [InlineData("1: a\n'1': b\n", false)]
    [InlineData("1: a\n1.0: b\n", false)]
    [InlineData("!x a: 1\n!y a: 2\n", false)]
    public void KeysAreDuplicatesWhenTheirValuesAreEqual(string yaml, bool duplicate) {
        if (duplicate) Assert.Equal(YamlErrorCodes.DuplicateKey, Assert.Throws<YamlException>(() => YamlDocument.Parse(yaml)).Code);
        else Assert.Equal(2, YamlDocument.Parse(yaml).Root!.AsMapping().Count);
    }

    private const string Merges = """
        base: &base {name: base, port: 80, tier: web}
        extra: &extra {port: 8080, debug: true}
        single: &single
          <<: *base
          name: own
        list:
          <<: [*extra, *base]
        nested:
          <<: *single
        """;

    [Fact]
    public void MergeKeysResolveInLookupsUnderYaml11WithoutChangingTheTree() {
        YamlDocument document = YamlDocument.Parse("%YAML 1.1\n---\n" + Merges + "\n");
        Assert.Equal(YamlSchema.Yaml11, document.Documents[0].Schema);
        YamlMapping root = document.Root!.AsMapping();

        YamlMapping single = root["single"].AsMapping();
        Assert.Equal(2, single.Count);
        Assert.Equal("own", single["name"].AsString());
        Assert.Equal(80, single["port"].GetValue<int>());

        // Earlier mappings in a merge list win over later ones.
        YamlMapping list = root["list"].AsMapping();
        Assert.Equal(8080, list["port"].GetValue<int>());
        Assert.Equal("web", list["tier"].AsString());

        // Merged mappings bring their own merges.
        Assert.Equal("own", document.Select("nested.name")!.AsString());
        Assert.Equal("web", document.Select("nested.tier")!.AsString());
        Assert.False(list.ContainsKey("missing"));
    }

    [Fact]
    public void MergeKeysAreOrdinaryKeysUnderTheCoreSchemaUnlessTheSchemaIsForced() {
        YamlMapping single = YamlDocument.Parse(Merges).Root!.AsMapping()["single"].AsMapping();
        Assert.False(single.ContainsKey("port"));
        Assert.True(single.ContainsKey("<<"));

        YamlDocument forced = YamlDocument.Parse(Merges, new YamlLoadOptions { Schema = YamlSchema.Yaml11 });
        Assert.Equal(80, forced.Select("single.port")!.GetValue<int>());
    }

    [Fact]
    public void RepeatedMergeKeysAreNotDuplicatesUnderYaml11() {
        YamlDocument document = YamlDocument.Parse("a: &a {x: 1}\nb: &b {y: 2}\nc:\n  <<: *a\n  <<: *b\n", new YamlLoadOptions { Schema = YamlSchema.Yaml11 });
        Assert.Equal(2, document.Select("c.y")!.GetValue<int>());
    }

    [Fact]
    public void EachDocumentResolvesThroughItsOwnSchema() {
        const string yaml = "%YAML 1.1\n---\nflag: yes\n...\n---\nflag: yes\n";
        YamlDocument document = YamlDocument.Parse(yaml);
        Assert.Equal("1.1", document.Documents[0].Version);
        Assert.True(document.Documents[0].Root.AsMapping()["flag"].GetValue<bool>());
        Assert.Null(document.Documents[1].Version);
        Assert.False(document.Documents[1].Root.AsMapping()["flag"].TryGetValue(out bool _));
        Assert.Equal("yes", document.Documents[1].Root.AsMapping()["flag"].GetValue<string>());

        YamlDocument forced = YamlDocument.Parse(yaml, new YamlLoadOptions { Schema = YamlSchema.Core });
        Assert.False(forced.Documents[0].Root.AsMapping()["flag"].TryGetValue(out bool _));
    }

    [Fact]
    public void TypedReadsFollowTheSchemaAndNameTheActualKind() {
        YamlDocument document = YamlDocument.Parse("big: 3000000000\nnone: ~\nseq: &s [1]\nalias: *s\nquoted: '12'\ncustom: !Ref name\n");
        YamlMapping root = document.Root!.AsMapping();
        Assert.Equal(3_000_000_000L, root["big"].GetValue<long>());
        Assert.False(root["big"].TryGetValue(out int _));
        Assert.Throws<FormatException>(() => root["big"].GetValue<int>());
        Assert.Null(root["none"].GetValue<int?>());
        Assert.Null(root["none"].GetValue<string>());
        Assert.Equal("~", root["none"].AsString());
        Assert.False(root["quoted"].TryGetValue(out int _));
        Assert.Equal("name", root["custom"].GetValue<string>());
        Assert.Equal("!Ref", root["custom"].Tag);
        Assert.Throws<NotSupportedException>(() => root["big"].GetValue<DateTime>());

        Assert.IsType<YamlAlias>(root["alias"]);
        Assert.Equal(1, root["alias"].AsSequence()[0].GetValue<int>());
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => root["alias"].AsMapping());
        Assert.Contains("a sequence (through alias *s)", error.Message);
        Assert.Contains("is a scalar", Assert.Throws<InvalidOperationException>(() => root["big"].AsSequence()).Message);
    }
}
