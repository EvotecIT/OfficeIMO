using OfficeIMO.Yaml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// The writer's main proof: every valid suite case and corpus file, replayed event by event through YamlWriter,
// re-parses to the same events. Known gaps belong in Exceptions with a reason; the list only shrinks.
public class YamlWriterConformanceTests {
    private static readonly IReadOnlyDictionary<string, string> Exceptions = new Dictionary<string, string>(StringComparer.Ordinal);

    private static readonly Lazy<YamlTestSuite> Suite = new(() => YamlTestSuite.Load());
    private static readonly Lazy<IReadOnlyList<YamlCorpusFile>> Corpus = new(() => YamlCorpus.Load());

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
    [MemberData(nameof(ValidCaseIds))]
    public void EmittedSuiteEventsReparseUnchanged(string id) {
        YamlTestSuiteCase testCase = Suite.Value.Cases.Single(c => c.Id == id);
        Check(id, new YamlConformanceChecks(new OfficeIMOYamlConformanceEngine()).CheckEmit(id, testCase.Input));
    }

    [Theory]
    [MemberData(nameof(CorpusPaths))]
    public void EmittedCorpusEventsReparseUnchanged(string path) {
        YamlCorpusFile file = Corpus.Value.Single(f => f.Path == path);
        Check(path, new YamlConformanceChecks(new OfficeIMOYamlConformanceEngine()).CheckEmit(path, file.Yaml));
    }

    private static void Check(string id, YamlCheckResult result) {
        if (Exceptions.TryGetValue(id, out string? reason)) {
            Assert.True(result.Outcome != YamlCheckOutcome.Pass, id + " now passes; remove it from the exception list (" + reason + ").");
            return;
        }

        Assert.True(result.Outcome == YamlCheckOutcome.Pass, result.ToString());
    }
}
