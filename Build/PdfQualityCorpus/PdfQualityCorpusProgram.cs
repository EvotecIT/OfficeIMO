using System.Text.Json;

namespace OfficeIMO.PdfQualityCorpus;

internal static class PdfQualityCorpusProgram {
    internal static async Task<int> RunAsync(string[] args) {
        try {
            if (args.Length == 0) throw new ArgumentException(Usage);
            if (string.Equals(args[0], "verify-markdown-contract", StringComparison.Ordinal)) {
                VerifyMarkdownContract();
                Console.WriteLine("PDF quality corpus Markdown contract passed.");
                return 0;
            }
            if (string.Equals(args[0], "verify-runner-contracts", StringComparison.Ordinal)) {
                VerifyRunnerContracts();
                Console.WriteLine("PDF quality corpus runner contracts passed.");
                return 0;
            }
            if (string.Equals(args[0], "probe", StringComparison.Ordinal)) {
                QualityCaseResult result = PdfQualityCorpusWorker.Probe(PdfQualityCorpusCommandLine.ParseProbe(args));
                Console.Write(JsonSerializer.Serialize(result, QualityJson.Options));
                return result.Outcome == "passed" ? 0 : 1;
            }
            if (!string.Equals(args[0], "run", StringComparison.Ordinal)) throw new ArgumentException(Usage);
            QualityRunOptions options = PdfQualityCorpusCommandLine.ParseRun(args);
            QualityReport report = await PdfQualityCorpusCoordinator.RunAsync(options).ConfigureAwait(false);
            PdfQualityCorpusReportWriter.Write(report, options.JsonReportPath, options.MarkdownReportPath);
            Console.WriteLine("PDF quality corpus: " + report.Totals.Passed + "/" + report.Totals.Cases + " passed");
            Console.WriteLine("Operational score: " + report.Totals.OperationalScore.ToString("P2", System.Globalization.CultureInfo.InvariantCulture));
            Console.WriteLine("Expectation score: " + report.Totals.ExpectationScore.ToString("P2", System.Globalization.CultureInfo.InvariantCulture));
            Console.WriteLine("JSON: " + options.JsonReportPath);
            Console.WriteLine("Markdown: " + options.MarkdownReportPath);
            return report.Totals.Failed == 0 && report.Totals.TimedOut == 0 ? 0 : 1;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception.Message);
            return 2;
        }
    }

    private static void VerifyMarkdownContract() {
        string unsafeValue = "<script>alert(1)</script> | [link](https://example.test) `code`\r\n## heading \u202E";
        string escaped = PdfQualityCorpusReportWriter.EscapeMarkdown(unsafeValue);
        string[] forbidden = { "<script", "</script>", "[link]", "(https://example.test)", "`code`", "\r", "\n", "\u202E" };
        foreach (string value in forbidden) {
            if (escaped.Contains(value, StringComparison.Ordinal)) {
                throw new InvalidOperationException("Markdown escaping left an unsafe token: " + value + ".");
            }
        }
        if (!escaped.Contains("&lt;script&gt;", StringComparison.Ordinal) ||
            !escaped.Contains("&#91;link&#93;", StringComparison.Ordinal) ||
            !escaped.Contains("&#96;code&#96;", StringComparison.Ordinal)) {
            throw new InvalidOperationException("Markdown escaping did not preserve safe visible text.");
        }
    }

    private static void VerifyRunnerContracts() {
        Expect<ArgumentException>(() => PdfQualityCorpusCommandLine.ParseRun(new[] { "run", "--unexpected", "value" }));
        var traversal = new QualityCase { Id = "traversal", File = Path.Combine("..", "escape.pdf") };
        Expect<InvalidDataException>(() => PdfQualityCorpusManifest.ResolveCasePath(Path.GetTempPath(), traversal));
        var rooted = new QualityCase { Id = "rooted", File = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "escape.pdf")) };
        Expect<InvalidDataException>(() => PdfQualityCorpusManifest.ResolveCasePath(Path.GetTempPath(), rooted));
        PdfQualityCorpusCoordinator.VerifyFailureScoringContract();
    }

    private static void Expect<TException>(Action action) where TException : Exception {
        try {
            action();
        } catch (TException) {
            return;
        }
        throw new InvalidOperationException("Expected " + typeof(TException).Name + " was not thrown.");
    }

    private const string Usage = "Usage: run --manifest <path> --root <directory> --json <path> --markdown <path> [--max-file-bytes <n>] [--max-render-pages <n>] [--timeout-seconds <n>] [--parallelism <n>] [--max-worker-memory-bytes <n>] | verify-markdown-contract | verify-runner-contracts";
}
