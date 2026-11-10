using OfficeIMO.Data.Conversions;
using OfficeIMO.Yaml;
using OfficeIMO.Yaml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Data.Conversions.Tests;

// Cross-format oracle: every yaml-test-suite case with in.json converts to equal JSON data through both the streaming
// route and the document-model route. Known gaps belong in Exceptions with a reason; the list only shrinks.
public class YamlJsonOracleTests {
    private static readonly IReadOnlyDictionary<string, string> Exceptions = new Dictionary<string, string>(StringComparer.Ordinal);

    private static readonly Lazy<YamlTestSuite> Suite = new(() => YamlTestSuite.Load());

    public static TheoryData<string> JsonCaseIds() {
        var data = new TheoryData<string>();
        foreach (YamlTestSuiteCase testCase in Suite.Value.Cases) {
            if (!testCase.IsError && testCase.Json != null) data.Add(testCase.Id);
        }

        return data;
    }

    [Theory]
    [MemberData(nameof(JsonCaseIds))]
    public void StreamingAndModelRoutesMatchInJson(string id) {
        YamlTestSuiteCase testCase = Suite.Value.Cases.Single(c => c.Id == id);
        IReadOnlyList<YamlConformanceNode> expected = YamlJsonReader.ReadStream(testCase.Json!);

        using var output = new MemoryStream();
        DataConversionEngine.Convert(new MemoryStream(testCase.Input), DataFormat.Yaml, output, DataFormat.Json, null);
        string? streaming = YamlJsonEquivalence.FindDifference(expected, Documents(output.ToArray(), expected.Count));

        using YamlDocument document = YamlDocument.Load(new MemoryStream(testCase.Input));
        using var modelOutput = new MemoryStream();
        document.SaveAsJson(modelOutput);
        string? model = YamlJsonEquivalence.FindDifference(expected, Documents(modelOutput.ToArray(), expected.Count));

        string? failure = streaming != null ? "streaming: " + streaming : model != null ? "model: " + model : null;
        if (Exceptions.TryGetValue(id, out string? reason)) {
            Assert.True(failure != null, id + " now passes; remove it from the exception list (" + reason + ").");
            return;
        }

        Assert.True(failure == null, id + " " + failure + "\n" + System.Text.Encoding.UTF8.GetString(output.ToArray()));
    }

    // YAML output re-reads to the suite's data: every in.json case goes YAML -> JSON -> YAML -> JSON through the writer.
    [Theory]
    [MemberData(nameof(JsonCaseIds))]
    public void YamlWrittenFromJsonReadsBackToInJson(string id) {
        YamlTestSuiteCase testCase = Suite.Value.Cases.Single(c => c.Id == id);
        IReadOnlyList<YamlConformanceNode> expected = YamlJsonReader.ReadStream(testCase.Json!);
        if (expected.Count > 1) return;
        using var json = new MemoryStream();
        DataConversionEngine.Convert(new MemoryStream(testCase.Input), DataFormat.Yaml, json, DataFormat.Json, null);
        using var yaml = new MemoryStream();
        DataConversionEngine.Convert(new MemoryStream(json.ToArray()), DataFormat.Json, yaml, DataFormat.Yaml, null);
        using var back = new MemoryStream();
        DataConversionEngine.Convert(new MemoryStream(yaml.ToArray()), DataFormat.Yaml, back, DataFormat.Json, null);
        string? difference = YamlJsonEquivalence.FindDifference(expected, Documents(back.ToArray(), expected.Count));
        Assert.True(difference == null, id + " " + difference + "\n" + System.Text.Encoding.UTF8.GetString(yaml.ToArray()));
    }

    // A %YAML 1.1 document resolves the same through the streaming reader and the loaded document: 1.1 booleans,
    // timestamps and merge keys alike.
    [Fact]
    public void StreamingAndModelRoutesAgreeOnYaml11Documents() {
        const string yaml = "%YAML 1.1\n---\nbase: &b {size: 1, shared: yes}\nderived:\n  <<: *b\n  size: 2\nwhen: 2001-12-14 21:59:43.10 -5\nflag: off\n";
        using var streamed = new MemoryStream();
        DataConversionEngine.Convert(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(yaml)), DataFormat.Yaml, streamed, DataFormat.Json, null);
        using YamlDocument document = YamlDocument.Parse(yaml);
        using var modelled = new MemoryStream();
        document.SaveAsJson(modelled);
        string json = System.Text.Encoding.UTF8.GetString(streamed.ToArray());
        Assert.Equal(json, System.Text.Encoding.UTF8.GetString(modelled.ToArray()));
        using System.Text.Json.JsonDocument parsed = System.Text.Json.JsonDocument.Parse(json);
        System.Text.Json.JsonElement root = parsed.RootElement;
        Assert.Equal(2, root.GetProperty("derived").GetProperty("size").GetInt32());
        Assert.True(root.GetProperty("derived").GetProperty("shared").GetBoolean());
        Assert.False(root.GetProperty("flag").GetBoolean());
        Assert.Equal("2001-12-14T21:59:43.10-05:00", root.GetProperty("when").GetString());
    }

    // One document is the JSON value itself; several are the items of one array; none is null.
    private static IReadOnlyList<YamlConformanceNode> Documents(byte[] json, int expectedCount) {
        YamlConformanceNode value = YamlJsonReader.Read(System.Text.Encoding.UTF8.GetString(json));
        if (expectedCount == 1) return new[] { value };
        if (expectedCount == 0 && value.Kind == YamlConformanceNodeKind.Null) return Array.Empty<YamlConformanceNode>();
        return value.Kind == YamlConformanceNodeKind.Sequence ? value.Items : new[] { value };
    }
}
