using System.Text;
using OfficeIMO.Yaml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// The model's editing proof: scripted edits on every corpus file and valid suite case must save text that re-parses to
// the edited model's data and comments, changes only the edited region, and is the same on every run.
public class YamlEditQualificationTests {
    private static readonly Lazy<YamlTestSuite> Suite = new(() => YamlTestSuite.Load());
    private static readonly Lazy<IReadOnlyList<YamlCorpusFile>> Corpus = new(() => YamlCorpus.Load());

    public static TheoryData<string> CorpusPaths() {
        var data = new TheoryData<string>();
        foreach (YamlCorpusFile file in Corpus.Value) data.Add(file.Path);
        return data;
    }

    public static TheoryData<string> ValidCaseIds() {
        var data = new TheoryData<string>();
        foreach (YamlTestSuiteCase testCase in Suite.Value.Cases) {
            if (!testCase.IsError) data.Add(testCase.Id);
        }
        return data;
    }

    [Theory]
    [MemberData(nameof(CorpusPaths))]
    public void ScriptedEditsOnCorpusFilesChangeOnlyTheEditedText(string path) {
        YamlCorpusFile file = Corpus.Value.Single(f => f.Path == path);
        Check(YamlEditScripts.Run(path, YamlTextTestHelpers.DecodeSuiteInput(file.Yaml), perKind: 12));
    }

    [Theory]
    [MemberData(nameof(ValidCaseIds))]
    public void ScriptedEditsOnSuiteCasesChangeOnlyTheEditedText(string id) {
        YamlTestSuiteCase testCase = Suite.Value.Cases.Single(c => c.Id == id);
        Check(YamlEditScripts.Run(id, YamlTextTestHelpers.DecodeSuiteInput(testCase.Input), perKind: 3));
    }

    private static void Check(IReadOnlyList<YamlEditResult> results) {
        var failures = results.Where(r => r.Failure is not null).ToList();
        Assert.True(failures.Count == 0, failures.Count + " of " + results.Count + " edits failed:\n" + YamlEditScripts.Describe(failures, 2));
    }
}
