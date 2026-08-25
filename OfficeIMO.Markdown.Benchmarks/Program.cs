using BenchmarkDotNet.Running;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--parse-evidence-probe", StringComparison.Ordinal)) {
    Environment.ExitCode = MarkdownParseEvidenceRunner.RunProbe(args[1..]);
    return;
}

if (args.Length > 0 && string.Equals(args[0], "--parse-evidence", StringComparison.Ordinal)) {
    Environment.ExitCode = MarkdownParseEvidenceRunner.RunEvidence(args[1..]);
    return;
}

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
