using OfficeIMO.Yaml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// The merge's proof on real base files with overlays in the style of real per-environment files: the merged file differs
// from the base only where the report says, its data equals an independent Helm-rules merge of YamlDotNet-loaded data,
// YamlDotNet and SharpYaml read it alike, and merging again writes the same bytes.
public class YamlMergeQualificationTests {
    private static readonly Lazy<IReadOnlyList<YamlMergePair>> Pairs = new(() => YamlMergeCorpus.Load());

    public static TheoryData<string> PairIds() {
        var data = new TheoryData<string>();
        foreach (YamlMergePair pair in Pairs.Value) data.Add(pair.Id);
        return data;
    }

    [Theory]
    [MemberData(nameof(PairIds))]
    public void MergedPairsAreMinimalDeterministicAndEqualToTheReferenceMerge(string id) {
        YamlMergePair pair = Pairs.Value.Single(p => p.Id == id);
        IReadOnlyList<string> failures = YamlMergeChecks.Run(pair.Base, pair.Overlay, pair.Sequences, pair.NullRemovesKey, out _);
        Assert.True(failures.Count == 0, string.Join("\n", failures));
    }
}
