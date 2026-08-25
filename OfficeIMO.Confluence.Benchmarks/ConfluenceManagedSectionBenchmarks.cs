using BenchmarkDotNet.Attributes;
using OfficeIMO.Confluence;

namespace OfficeIMO.Confluence.Benchmarks;

[MemoryDiagnoser]
[ShortRunJob]
public class ConfluenceManagedSectionBenchmarks {
    private string _existingBody = string.Empty;
    private string _replacement = string.Empty;

    [Params(16 * 1024, 1024 * 1024)]
    public int PageCharacters { get; set; }

    [GlobalSetup]
    public void Setup() {
        ConfluenceManagedSectionCorpus corpus = ConfluenceManagedSectionCorpusFactory.Create(PageCharacters);
        _existingBody = corpus.ExistingBody;
        _replacement = corpus.Replacement;

        ConfluenceManagedSectionResult result = ReplaceSection();
        if (!result.Changed || result.WasCreated || !result.UpdatedBody.Contains(_replacement, StringComparison.Ordinal)) {
            throw new InvalidOperationException("Managed-section benchmark validation failed.");
        }
        if (string.Equals(result.OriginalSha256, result.UpdatedSha256, StringComparison.Ordinal)) {
            throw new InvalidOperationException("Managed-section benchmark hashes did not change.");
        }
    }

    [Benchmark]
    public ConfluenceManagedSectionResult ReplaceSection() =>
        ConfluenceManagedSection.Apply(_existingBody, ConfluenceManagedSectionCorpusFactory.SectionId, _replacement);
}
