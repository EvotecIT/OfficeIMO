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

    // One document is the JSON value itself; several are the items of one array; none is null.
    private static IReadOnlyList<YamlConformanceNode> Documents(byte[] json, int expectedCount) {
        YamlConformanceNode value = YamlJsonReader.Read(System.Text.Encoding.UTF8.GetString(json));
        if (expectedCount == 1) return new[] { value };
        if (expectedCount == 0 && value.Kind == YamlConformanceNodeKind.Null) return Array.Empty<YamlConformanceNode>();
        return value.Kind == YamlConformanceNodeKind.Sequence ? value.Items : new[] { value };
    }
}
