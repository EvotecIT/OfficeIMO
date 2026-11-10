using Xunit;

namespace OfficeIMO.Yaml.Tests;

// Editing contracts the scripted corpus edits do not reach: exact output for the documented examples, value equality
// keeping source spellings, layout inference for new nodes, path editing, anchors and the shared-node diagnostic.
public class YamlDocumentEditingTests {
    [Fact]
    public void EqualValuesKeepTheirSourceSpelling() {
        YamlDocument document = YamlDocument.Parse("debug: TRUE\nport: 0x1F\nname: 'web'\n");
        YamlMapping root = document.Root!.AsMapping();
        root["debug"] = true;
        root["port"] = 31;
        root["name"] = "web";
        Assert.Equal("debug: TRUE\nport: 0x1F\nname: 'web'\n", document.ToString());
        root["port"] = 8080;
        Assert.Equal("debug: TRUE\nport: 8080\nname: 'web'\n", document.ToString());
    }

    [Fact]
    public void ChangedStringsKeepTheirStyleWhenItFitsAndAreQuotedWhenAmbiguous() {
        YamlDocument document = YamlDocument.Parse("a: plain # note\nb: \"double\"\nc: |\n  line\n");
        YamlMapping root = document.Root!.AsMapping();
        root["a"] = "yes";
        root["b"] = "other";
        root["c"] = "one\ntwo\n";
        Assert.Equal("a: \"yes\" # note\nb: \"other\"\nc: |\n  one\n  two\n", document.ToString());
    }

    [Fact]
    public void NewEntriesFollowTheDocumentsIndentationAndSequenceStyle() {
        YamlDocument document = YamlDocument.Parse("spec:\n    containers:\n      - name: web\n");
        YamlMapping spec = document.Select("spec")!.AsMapping();
        spec.Add("volumes", new YamlSequence { new YamlMapping { { "name", "data" } } });
        spec.Add("selector", new YamlMapping { { "app", "web" } });
        Assert.Equal(
            "spec:\n    containers:\n      - name: web\n    volumes:\n      - name: data\n    selector:\n        app: web\n",
            document.ToString());
    }

    [Fact]
    public void RemovingAnEntryTakesItsCommentsAndLeavesNeighborsAlone() {
        YamlDocument document = YamlDocument.Parse("a: 1 # about a\n\n# about b\nb: 2 # trailing b\n# about c\nc: 3\n");
        document.Root!.AsMapping().Remove("b");
        Assert.Equal("a: 1 # about a\n# about c\nc: 3\n", document.ToString());
    }

    [Fact]
    public void RemovingAnEntryKeepsACommentBlockThatABlankLineSeparatesFromIt() {
        // The shape of Helm values: a commented-out example under one key, then the next key's own note.
        const string yaml = "affinity: {}\n#   nodeAffinity:\n#     required: true\n\n# Annotations to be added to pods\npodAnnotations:\n  # Fix for slow upgrades\n  safe-to-evict: \"true\"\n\n# Extra labels\npodLabels: {}\n";
        YamlDocument document = YamlDocument.Parse(yaml);
        document.Root!.AsMapping().Remove("podAnnotations");
        Assert.Equal("affinity: {}\n#   nodeAffinity:\n#     required: true\n\n# Extra labels\npodLabels: {}\n", document.ToString());
    }

    [Fact]
    public void PathSetCreatesMissingMappingsAndAddAppendsToSequences() {
        using YamlDocument document = YamlDocument.Create();
        document.Set("metadata.name", "web");
        document.Set("spec.replicas", 3);
        document.Add("spec.ports", 80);
        document.Add("spec.ports", 443);
        document.Set("spec.ports[2]", 8080);
        Assert.Equal("metadata:\n  name: web\nspec:\n  replicas: 3\n  ports:\n  - 80\n  - 443\n  - 8080\n", document.ToString());
        Assert.True(document.Remove("spec.ports[0]"));
        Assert.False(document.Remove("spec.missing"));
        Assert.Throws<InvalidOperationException>(() => document.Set("metadata.name.first", "x"));
        Assert.Throws<InvalidOperationException>(() => document.Remove("metadata.name.first"));
        Assert.Equal("$[1].spec.ports[1]", YamlDocument.Parse("a: 1\n---\nspec:\n  ports: [1, 2]\n").Select("$[1].spec.ports[1]")!.GetPath());
    }

    [Fact]
    public void ChainedMutatorsSetCommentsAndBlankLines() {
        YamlDocument document = YamlDocument.Parse("image: nginx\n");
        YamlMapping root = document.Root!.AsMapping();
        root.Add("tag", "1.25").SetComment("pinned by CI").SetCommentBefore("Release").SetSpaceBefore(true);
        Assert.Equal("image: nginx\n\n# Release\ntag: '1.25' # pinned by CI\n", document.ToString());
    }

    [Fact]
    public void DocumentHeaderAndFooterAreEditable() {
        YamlDocument document = YamlDocument.Parse("# old license\n\na: 1\n# old footer\n");
        document.Documents[0].SetCommentBefore("SPDX-License-Identifier: MIT").SetCommentAfter(null);
        Assert.Equal("# SPDX-License-Identifier: MIT\n\na: 1\n", document.ToString());
    }

    [Fact]
    public void RenamingAnAnchorRenamesItsAliasesAndRemovingAnAliasedNodeFails() {
        YamlDocument document = YamlDocument.Parse("base: &defaults\n  timeout: 30\nprod:\n  <<: *defaults\n");
        document.Select("base")!.SetAnchor("common");
        Assert.Equal("base: &common\n  timeout: 30\nprod:\n  <<: *common\n", document.ToString());
        Assert.Throws<InvalidOperationException>(() => document.Root!.AsMapping().Remove("base"));
    }

    [Fact]
    public void EditingAnAliasedNodeRecordsOneDiagnostic() {
        YamlDocument document = YamlDocument.Parse("base: &b\n  timeout: 30\nprod: *b\n");
        document.Select("prod")!.AsMapping()["timeout"] = 60;
        document.Select("prod")!.AsMapping()["retries"] = 2;
        OfficeDiagnostic diagnostic = Assert.Single(document.Diagnostics);
        Assert.Equal(OfficeDiagnosticCodes.YamlAliasedNodeEdited, diagnostic.Code);
        Assert.Equal("base: &b\n  timeout: 60\n  retries: 2\nprod: *b\n", document.ToString());
    }

    [Fact]
    public void LineEndingsFollowTheSource() {
        YamlDocument document = YamlDocument.Parse("a: 1\r\nb:\r\n  c: 2\r\n");
        document.Select("b")!.AsMapping().Add("d", new YamlSequence { "x" });
        Assert.Equal("a: 1\r\nb:\r\n  c: 2\r\n  d:\r\n  - x\r\n", document.ToString());
    }
}
