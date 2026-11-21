using OfficeIMO.Yaml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// The model's main proof: every suite case with in.json loads to equal values, and every valid case and corpus file
// saves back to its exact input bytes. Known gaps belong in Exceptions with a reason; the list only shrinks.
public class YamlDocumentConformanceTests {
    /// <summary>Cases the model is known to get wrong, with the reason. A listed case that starts passing fails the test.</summary>
    private static readonly IReadOnlyDictionary<string, string> Exceptions = new Dictionary<string, string>(StringComparer.Ordinal);

    private static readonly Lazy<YamlTestSuite> Suite = new(() => YamlTestSuite.Load());
    private static readonly Lazy<IReadOnlyList<YamlCorpusFile>> Corpus = new(() => YamlCorpus.Load());

    public static TheoryData<string> JsonCaseIds() {
        var data = new TheoryData<string>();
        foreach (YamlTestSuiteCase testCase in Suite.Value.Cases) {
            if (!testCase.IsError && testCase.Json != null) data.Add(testCase.Id);
        }
        return data;
    }

    public static TheoryData<string> ValidCaseIds() {
        var data = new TheoryData<string>();
        foreach (YamlTestSuiteCase testCase in Suite.Value.Cases) {
            if (!testCase.IsError) data.Add(testCase.Id);
        }
        return data;
    }

    public static TheoryData<string> CorpusPaths() {
        var data = new TheoryData<string>();
        foreach (YamlCorpusFile file in Corpus.Value) data.Add(file.Path);
        return data;
    }

    [Theory]
    [MemberData(nameof(JsonCaseIds))]
    public void LoadedValuesMatchInJson(string id) {
        YamlTestSuiteCase testCase = Suite.Value.Cases.Single(c => c.Id == id);
        Check(id, new YamlConformanceChecks(new OfficeIMOYamlConformanceEngine()).CheckJson(testCase));
    }

    [Theory]
    [MemberData(nameof(ValidCaseIds))]
    public void UnchangedSaveReproducesSuiteInputBytes(string id) {
        YamlTestSuiteCase testCase = Suite.Value.Cases.Single(c => c.Id == id);
        Check(id, new YamlConformanceChecks(new OfficeIMOYamlConformanceEngine()).CheckRoundTrip(id, testCase.Input));
    }

    [Theory]
    [MemberData(nameof(CorpusPaths))]
    public void UnchangedSaveReproducesCorpusFileBytes(string path) {
        YamlCorpusFile file = Corpus.Value.Single(f => f.Path == path);
        Check(path, new YamlConformanceChecks(new OfficeIMOYamlConformanceEngine()).CheckRoundTrip(path, file.Yaml));
    }

    private static void Check(string id, YamlCheckResult result) {
        if (Exceptions.TryGetValue(id, out string? reason)) {
            Assert.True(result.Outcome != YamlCheckOutcome.Pass, id + " now passes; remove it from the exception list (" + reason + ").");
            return;
        }

        Assert.True(result.Outcome == YamlCheckOutcome.Pass, result.ToString());
    }
}
