using BenchmarkDotNet.Running;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Benchmarks;

if (args.Length == 1 && string.Equals(args[0], "--validate-equivalence", StringComparison.Ordinal)) {
    foreach (string corpusName in MarkdownBenchmarkCorpus.Names) {
        MarkdownBenchmarkValidation.AssertCommonMarkEquivalent(
            corpusName,
            MarkdownBenchmarkCorpus.Get(corpusName),
            MarkdownReaderOptions.CreateCommonMarkProfile(),
            MarkdownBenchmarkValidation.CreateOfficeCommonMarkHtmlOptions());
    }

    HtmlToMarkdownBenchmarkValidation.AssertAllDefaultComparisons();
    Console.WriteLine("Benchmark equivalence validation passed.");
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
