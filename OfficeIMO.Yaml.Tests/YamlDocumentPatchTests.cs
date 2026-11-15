using Xunit;

namespace OfficeIMO.Yaml.Tests;

// RFC 6902 JSON Patch and RFC 7386 JSON Merge Patch on the lossless model: the RFC examples by value, then the YAML
// contract: untouched lines stay byte for byte, a failed patch changes nothing, and aliases and merge keys follow the
// typed view without editing merged mappings.
public class YamlDocumentPatchTests {
    private static string Save(YamlDocument document) {
        var stream = new MemoryStream();
        document.Save(stream);
        return new System.Text.UTF8Encoding(false).GetString(stream.ToArray());
    }

    /// <summary>A JSON-like rendering with sorted keys, so documents compare by value.</summary>
    private static string Value(YamlNode? node) {
        if (node is null) return "<none>";
        switch (node.Resolved) {
            case YamlMapping mapping:
                var entries = new List<string>();
                foreach (YamlMappingEntry entry in mapping) entries.Add(entry.Key.AsString() + ":" + Value(entry.Value));
                entries.Sort(StringComparer.Ordinal);
                return "{" + string.Join(",", entries) + "}";
            case YamlSequence sequence:
                var items = new List<string>();
                foreach (YamlNode item in sequence) items.Add(Value(item));
                return "[" + string.Join(",", items) + "]";
            default:
                var scalar = (YamlScalar)node.Resolved;
                return scalar.Kind switch {
                    YamlScalarKind.Null => "null",
                    YamlScalarKind.Integer or YamlScalarKind.Float => "n" + scalar.GetValue<double>().ToString(System.Globalization.CultureInfo.InvariantCulture),
                    YamlScalarKind.Boolean => scalar.GetValue<bool>() ? "true" : "false",
                    _ => "'" + scalar.Value + "'"
                };
        }
    }

    private static void AssertValue(string expectedJson, YamlDocument document) {
        Assert.Equal(Value(YamlDocument.Parse(expectedJson).Root), Value(document.Root));
        // The saved text reads back to the same value.
        Assert.Equal(Value(YamlDocument.Parse(expectedJson).Root), Value(YamlDocument.Parse(Save(document)).Root));
    }

    // ------------------------------------------------------------------ RFC 6902 appendix A

    [Theory]
    [InlineData("A.1", """{ "foo": "bar" }""", """[ { "op": "add", "path": "/baz", "value": "qux" } ]""", """{ "baz": "qux", "foo": "bar" }""")]
    [InlineData("A.2", """{ "foo": [ "bar", "baz" ] }""", """[ { "op": "add", "path": "/foo/1", "value": "qux" } ]""", """{ "foo": [ "bar", "qux", "baz" ] }""")]
    [InlineData("A.3", """{ "baz": "qux", "foo": "bar" }""", """[ { "op": "remove", "path": "/baz" } ]""", """{ "foo": "bar" }""")]
    [InlineData("A.4", """{ "foo": [ "bar", "qux", "baz" ] }""", """[ { "op": "remove", "path": "/foo/1" } ]""", """{ "foo": [ "bar", "baz" ] }""")]
    [InlineData("A.5", """{ "baz": "qux", "foo": "bar" }""", """[ { "op": "replace", "path": "/baz", "value": "boo" } ]""", """{ "baz": "boo", "foo": "bar" }""")]
    [InlineData("A.6", """{ "foo": { "bar": "baz", "waldo": "fred" }, "qux": { "corge": "grault" } }""", """[ { "op": "move", "from": "/foo/waldo", "path": "/qux/thud" } ]""", """{ "foo": { "bar": "baz" }, "qux": { "corge": "grault", "thud": "fred" } }""")]
    [InlineData("A.7", """{ "foo": [ "all", "grass", "cows", "eat" ] }""", """[ { "op": "move", "from": "/foo/1", "path": "/foo/3" } ]""", """{ "foo": [ "all", "cows", "eat", "grass" ] }""")]
    [InlineData("A.8", """{ "baz": "qux", "foo": [ "a", 2, "c" ] }""", """[ { "op": "test", "path": "/baz", "value": "qux" }, { "op": "test", "path": "/foo/1", "value": 2 } ]""", """{ "baz": "qux", "foo": [ "a", 2, "c" ] }""")]
    [InlineData("A.10", """{ "foo": "bar" }""", """[ { "op": "add", "path": "/child", "value": { "grandchild": { } } } ]""", """{ "foo": "bar", "child": { "grandchild": { } } }""")]
    [InlineData("A.11", """{ "foo": "bar" }""", """[ { "op": "add", "path": "/baz", "value": "qux", "xyz": 123 } ]""", """{ "foo": "bar", "baz": "qux" }""")]
    [InlineData("A.14", """{ "/": 9, "~1": 10 }""", """[ { "op": "test", "path": "/~01", "value": 10 } ]""", """{ "/": 9, "~1": 10 }""")]
    [InlineData("A.16", """{ "foo": ["bar"] }""", """[ { "op": "add", "path": "/foo/-", "value": ["abc", "def"] } ]""", """{ "foo": ["bar", ["abc", "def"]] }""")]
    public void AppliesTheRfc6902Examples(string section, string target, string patch, string expected) {
        _ = section;
        using YamlDocument document = YamlDocument.Parse(target);
        document.ApplyJsonPatch(patch);
        AssertValue(expected, document);
    }

    [Theory]
    [InlineData("A.9", """{ "baz": "qux" }""", """[ { "op": "test", "path": "/baz", "value": "bar" } ]""")]
    [InlineData("A.12", """{ "foo": "bar" }""", """[ { "op": "add", "path": "/baz/bat", "value": "qux" } ]""")]
    [InlineData("A.15", """{ "/": 9, "~1": 10 }""", """[ { "op": "test", "path": "/~01", "value": "10" } ]""")]
    public void FailsTheRfc6902ErrorExamples(string section, string target, string patch) {
        _ = section;
        using YamlDocument document = YamlDocument.Parse(target);
        YamlPatchException exception = Assert.Throws<YamlPatchException>(() => document.ApplyJsonPatch(patch));
        Assert.Equal(0, exception.OperationIndex);
        Assert.Equal(target, Save(document));
    }

    [Fact]
    public void RejectsAnOperationWithARepeatedMemberAsInRfc6902A13() {
        using YamlDocument document = YamlDocument.Parse("""{ "baz": "qux" }""");
        Assert.ThrowsAny<YamlException>(() => document.ApplyJsonPatch("""[ { "op": "add", "path": "/baz", "value": "qux", "op": "remove" } ]"""));
    }

    // ------------------------------------------------------------------ atomic failure

    [Fact]
    public void AFailedTestLeavesTheDocumentUnchanged() {
        const string text = "# service\nname: web # the name\nreplicas: 2\nports:\n  - 80\n";
        using YamlDocument document = YamlDocument.Parse(text);

        YamlPatchException exception = Assert.Throws<YamlPatchException>(() => document.ApplyJsonPatch("""
            [
              { "op": "replace", "path": "/replicas", "value": 5 },
              { "op": "add", "path": "/ports/-", "value": 443 },
              { "op": "test", "path": "/name", "value": "api" }
            ]
            """));

        Assert.Equal(2, exception.OperationIndex);
        Assert.Equal("/name", exception.Path);
        Assert.Equal(text, Save(document));
        Assert.Equal(2, document.Select("replicas")!.GetValue<int>());
        Assert.False(document.Root!.IsModified);
        Assert.Empty(document.Diagnostics);
    }

    [Theory]
    [InlineData("""[ { "op": "add", "path": "replicas", "value": 1 } ]""", 0)]
    [InlineData("""[ { "op": "replace", "path": "/replicas", "value": 1 }, { "op": "remove", "path": "/ports/~2" } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/replicas", "value": 1 }, { "op": "remove", "path": "/ports/01" } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/replicas", "value": 1 }, { "op": "replace", "path": "/ports/1", "value": 1 } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/replicas", "value": 1 }, { "op": "remove", "path": "/ports/-" } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/replicas", "value": 1 }, { "op": "replace", "path": "/missing", "value": 1 } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/replicas", "value": 1 }, { "op": "move", "from": "/ports", "path": "/ports/0" } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/replicas", "value": 1 }, { "op": "add", "path": "/name/first", "value": 1 } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/replicas", "value": 1 }, { "op": "copy", "from": "/nothing", "path": "/x" } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/replicas", "value": 1 }, { "op": "remove", "path": "" } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/replicas", "value": 1 }, { "op": "rename", "path": "/name" } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/replicas", "value": 1 }, { "op": "replace", "path": "/name" } ]""", 1)]
    [InlineData("""[ { "op": "replace", "path": "/replicas", "value": 1 }, { "op": "move", "path": "/name" } ]""", 1)]
    [InlineData("""{ "op": "remove", "path": "/name" }""", -1)]
    public void InvalidPointersAndOperationsFailTheWholePatch(string patch, int failedOperation) {
        const string text = "name: web\nreplicas: 2\nports:\n  - 80\n";
        using YamlDocument document = YamlDocument.Parse(text);

        YamlPatchException exception = Assert.Throws<YamlPatchException>(() => document.ApplyJsonPatch(patch));

        Assert.Equal(failedOperation, exception.OperationIndex);
        Assert.Equal(text, Save(document));
    }

    // ------------------------------------------------------------------ layout

    [Fact]
    public void APatchChangesOnlyTheLinesItTouches() {
        const string text = """
            # Deployment for the web tier.
            apiVersion: apps/v1
            kind: Deployment
            metadata:
              name: 'web'   # keep the quotes
              labels:
                app: web
                tier: "frontend"
            spec:
              replicas: 2 # scaled by hand

              template:
                spec:
                  containers:
                    - name: web
                      image: nginx:1.25
                      ports: [80, 443]
            """;
        using YamlDocument document = YamlDocument.Parse(text);

        YamlPatchReport report = document.ApplyJsonPatch("""
            [
              { "op": "replace", "path": "/spec/replicas", "value": 3 },
              { "op": "remove", "path": "/metadata/labels/tier" },
              { "op": "add", "path": "/metadata/labels/env", "value": "prod" },
              { "op": "add", "path": "/spec/template/spec/containers/0/env", "value": [ { "name": "MODE", "value": "fast" } ] },
              { "op": "add", "path": "/spec/template/spec/containers/0/ports/-", "value": 8080 }
            ]
            """);

        Assert.Equal("""
            # Deployment for the web tier.
            apiVersion: apps/v1
            kind: Deployment
            metadata:
              name: 'web'   # keep the quotes
              labels:
                app: web
                env: prod
            spec:
              replicas: 3 # scaled by hand

              template:
                spec:
                  containers:
                    - name: web
                      image: nginx:1.25
                      ports: [80, 443, 8080]
                      env:
                        - name: MODE
                          value: fast

            """, Save(document));
        Assert.Equal(new[] { "metadata.labels.env", "spec.template.spec.containers[0].env", "spec.template.spec.containers[0].ports[2]" }, report.Added);
        Assert.Equal(new[] { "spec.replicas" }, report.Replaced);
        Assert.Equal(new[] { "metadata.labels.tier" }, report.Removed);
        Assert.Empty(report.Diagnostics);
    }

    [Fact]
    public void MoveKeepsTheMovedNodesCommentAndSpelling() {
        const string text = "old:\n  count: 0x1F # hex\n  keep: 1\nnew:\n  other: 2\n";
        using YamlDocument document = YamlDocument.Parse(text);

        YamlPatchReport report = document.ApplyJsonPatch("""[ { "op": "move", "from": "/old/count", "path": "/new/count" } ]""");

        Assert.Equal("old:\n  keep: 1\nnew:\n  other: 2\n  count: 0x1F # hex\n", Save(document));
        Assert.Equal(new[] { "old.count" }, report.Removed);
        Assert.Equal(new[] { "new.count" }, report.Added);
    }

    [Fact]
    public void APatchWrittenInYamlMeansTheSameAsJson() {
        using YamlDocument document = YamlDocument.Parse("name: web\n");
        using YamlDocument patch = YamlDocument.Parse("""
            - op: add
              path: /ports
              value: [80, 443]
            - op: test
              path: /ports/1
              value: 443.0
            """);

        document.ApplyJsonPatch(patch);

        Assert.Equal("name: web\nports:\n- 80\n- 443\n", Save(document));
    }

    [Fact]
    public void TheEmptyPointerAddressesTheRoot() {
        using YamlDocument empty = YamlDocument.Create();
        YamlPatchReport report = empty.ApplyJsonPatch("""[ { "op": "add", "path": "", "value": { "name": "web" } } ]""");
        Assert.Equal("name: web\n", Save(empty));
        Assert.Equal(new[] { "$" }, report.Added);

        using YamlDocument document = YamlDocument.Parse("a: 1\n");
        report = document.ApplyJsonPatch("""[ { "op": "replace", "path": "", "value": [1, 2] } ]""");
        Assert.Equal("- 1\n- 2\n", Save(document));
        Assert.Equal(new[] { "$" }, report.Replaced);
    }

    // ------------------------------------------------------------------ aliases and merge keys

    [Fact]
    public void PointersFollowAliasesAndEditTheAnchoredNode() {
        const string text = "defaults: &defaults\n  retries: 3\njob: *defaults\n";
        using YamlDocument document = YamlDocument.Parse(text);

        YamlPatchReport report = document.ApplyJsonPatch("""[ { "op": "replace", "path": "/job/retries", "value": 5 } ]""");

        Assert.Equal("defaults: &defaults\n  retries: 5\njob: *defaults\n", Save(document));
        Assert.Equal(new[] { "defaults.retries" }, report.Replaced);
        Assert.Equal(OfficeDiagnosticCodes.YamlAliasedNodeEdited, Assert.Single(report.Diagnostics).Code);
    }

    [Fact]
    public void KeysFromMergeKeysAreReadAndOverriddenButNeverRemoved() {
        const string text = "base: &base\n  retries: 3\n  timeout: 10\njob:\n  <<: *base\n  name: build\n";
        using YamlDocument document = YamlDocument.Parse(text);

        document.ApplyJsonPatch("""
            [
              { "op": "test", "path": "/job", "value": { "name": "build", "retries": 3, "timeout": 10 } },
              { "op": "replace", "path": "/job/retries", "value": 5 }
            ]
            """);
        Assert.Equal(text + "  retries: 5\n", Save(document));

        YamlPatchException exception = Assert.Throws<YamlPatchException>(() => document.ApplyJsonPatch("""[ { "op": "remove", "path": "/job/timeout" } ]"""));
        Assert.Contains("<< merge key", exception.Message);
        Assert.Equal(text + "  retries: 5\n", Save(document));
    }

    [Fact]
    public void RemovingOrMovingAnAnchorThatAliasesUseFails() {
        const string text = "defaults: &defaults\n  retries: 3\njob: *defaults\n";
        using YamlDocument document = YamlDocument.Parse(text);

        Assert.Throws<YamlPatchException>(() => document.ApplyJsonPatch("""[ { "op": "remove", "path": "/defaults" } ]"""));
        Assert.Throws<YamlPatchException>(() => document.ApplyJsonPatch("""[ { "op": "replace", "path": "/defaults", "value": 1 } ]"""));
        Assert.Throws<YamlPatchException>(() => document.ApplyJsonPatch("""[ { "op": "move", "from": "/job", "path": "/first" }, { "op": "move", "from": "/defaults", "path": "/last" } ]"""));
        Assert.Equal(text, Save(document));
    }

    [Fact]
    public void CopiedValuesExpandAliasesAndMerges() {
        using YamlDocument document = YamlDocument.Parse("base: &base\n  retries: 3\njob:\n  <<: *base\n  name: build\n");

        document.ApplyJsonPatch("""[ { "op": "copy", "from": "/job", "path": "/next" } ]""");

        Assert.Equal("base: &base\n  retries: 3\njob:\n  <<: *base\n  name: build\nnext:\n  name: build\n  retries: 3\n", Save(document));
    }

    // ------------------------------------------------------------------ RFC 7386 merge patch

    [Theory]
    [InlineData("""{"a":"b"}""", """{"a":"c"}""", """{"a":"c"}""")]
    [InlineData("""{"a":"b"}""", """{"b":"c"}""", """{"a":"b","b":"c"}""")]
    [InlineData("""{"a":"b"}""", """{"a":null}""", """{}""")]
    [InlineData("""{"a":"b","b":"c"}""", """{"a":null}""", """{"b":"c"}""")]
    [InlineData("""{"a":["b"]}""", """{"a":"c"}""", """{"a":"c"}""")]
    [InlineData("""{"a":"c"}""", """{"a":["b"]}""", """{"a":["b"]}""")]
    [InlineData("""{"a":{"b":"c"}}""", """{"a":{"b":"d","c":null}}""", """{"a":{"b":"d"}}""")]
    [InlineData("""{"a":[{"b":"c"}]}""", """{"a":[1]}""", """{"a":[1]}""")]
    [InlineData("""["a","b"]""", """["c","d"]""", """["c","d"]""")]
    [InlineData("""{"a":"b"}""", """["c"]""", """["c"]""")]
    [InlineData("""{"a":"foo"}""", """null""", """null""")]
    [InlineData("""{"a":"foo"}""", "\"bar\"", "\"bar\"")]
    [InlineData("""{"e":null}""", """{"a":1}""", """{"e":null,"a":1}""")]
    [InlineData("""[1,2]""", """{"a":"b","c":null}""", """{"a":"b"}""")]
    [InlineData("""{}""", """{"a":{"bb":{"ccc":null}}}""", """{"a":{"bb":{}}}""")]
    [InlineData("""{"a":"b"}""", """{"a":[null, {"b": null}]}""", """{"a":[null, {"b": null}]}""")]
    public void AppliesTheRfc7386Examples(string target, string patch, string expected) {
        using YamlDocument document = YamlDocument.Parse(target);
        document.ApplyMergePatch(patch);
        AssertValue(expected, document);
    }

    [Fact]
    public void AMergePatchKeepsTheDocumentsLayout() {
        const string text = """
            # Chart values.
            image:
              repository: nginx
              tag: "1.25" # pinned
            resources:
              limits:
                cpu: 1
            podAnnotations: {}
            replicas: 2
            """;
        using YamlDocument document = YamlDocument.Parse(text);

        YamlPatchReport report = document.ApplyMergePatch("""{ "image": { "tag": "1.27" }, "resources": null, "podAnnotations": { "team": "web" }, "service": { "port": 80, "debug": null } }""");

        Assert.Equal("""
            # Chart values.
            image:
              repository: nginx
              tag: "1.27" # pinned
            podAnnotations:
              team: web
            replicas: 2
            service:
              port: 80

            """, Save(document));
        Assert.Equal(new[] { "image.tag" }, report.Replaced);
        Assert.Equal(new[] { "resources" }, report.Removed);
        Assert.Equal(new[] { "podAnnotations.team", "service" }, report.Added);
        Assert.Empty(report.Diagnostics);
    }

    [Fact]
    public void AMergePatchNullForAMergedKeyFailsAndChangesNothing() {
        const string text = "base: &base\n  retries: 3\njob:\n  <<: *base\n  name: build\n";
        using YamlDocument document = YamlDocument.Parse(text);

        Assert.Throws<YamlPatchException>(() => document.ApplyMergePatch("""{ "job": { "name": "test", "retries": null } }"""));

        Assert.Equal(text, Save(document));
    }

    [Fact]
    public void AMergePatchInYamlKeepsItsComments() {
        using YamlDocument document = YamlDocument.Parse("name: web\n");
        using YamlDocument patch = YamlDocument.Parse("name: web\n# Scaled for launch.\nreplicas: 4 # four\n");

        document.ApplyMergePatch(patch);

        Assert.Equal("name: web\n# Scaled for launch.\nreplicas: 4 # four\n", Save(document));
    }

    [Fact]
    public void AnEmptyMergePatchChangesNothing() {
        using YamlDocument document = YamlDocument.Parse("name: web\n");
        YamlPatchReport report = document.ApplyMergePatch("");
        Assert.Equal("name: web\n", Save(document));
        Assert.Empty(report.Added);
        Assert.Empty(report.Replaced);
        Assert.Empty(report.Removed);
    }
}
