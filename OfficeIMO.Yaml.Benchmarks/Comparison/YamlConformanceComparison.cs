using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using OfficeIMO.Yaml.Benchmarks.Adapters;
using OfficeIMO.Yaml.Tests.Conformance;

namespace OfficeIMO.Yaml.Benchmarks.Comparison;

/// <summary>
/// Runs the reference libraries through the conformance harness and the corpus and writes a compact Markdown report.
/// Usage: <c>--conformance [--output report.md] [--details] [--timeout seconds] [--python path]</c>.
/// </summary>
internal static class YamlConformanceComparison {
    private static readonly YamlConformanceCheck[] SuiteChecks = {
        YamlConformanceCheck.Events, YamlConformanceCheck.EventContent, YamlConformanceCheck.ErrorCase, YamlConformanceCheck.Json,
        YamlConformanceCheck.RoundTrip, YamlConformanceCheck.RoundTripContent, YamlConformanceCheck.Emit
    };

    private static readonly YamlConformanceCheck[] CorpusChecks = {
        YamlConformanceCheck.Accepts, YamlConformanceCheck.RoundTrip, YamlConformanceCheck.RoundTripContent, YamlConformanceCheck.Emit
    };

    public static int Run(string[] args) {
        string? outputPath = Option(args, "--output");
        bool details = args.Any(arg => string.Equals(arg, "--details", StringComparison.OrdinalIgnoreCase));
        double seconds = double.Parse(Option(args, "--timeout") ?? "10", CultureInfo.InvariantCulture);

        YamlTestSuite suite = YamlTestSuite.Load();
        IReadOnlyList<YamlCorpusFile> corpus = YamlCorpus.Load();
        IYamlConformanceEngine[] engines = { new OfficeIMOYamlConformanceEngine(), new YamlDotNetConformanceEngine(), new SharpYamlConformanceEngine(), new SharpYamlSyntaxConformanceEngine(), new VYamlConformanceEngine() };

        var runs = new List<EngineRun>();
        foreach (IYamlConformanceEngine engine in engines) {
            var checks = new YamlConformanceChecks(engine, TimeSpan.FromSeconds(seconds));
            var suiteResults = suite.Cases.SelectMany(checks.CheckSuiteCase).Concat(suite.Cases.Select(checks.CheckEventContent)).ToList();
            var corpusResults = corpus.SelectMany(checks.CheckCorpusFile).ToList();
            runs.Add(new EngineRun(engine, suiteResults, corpusResults));
            Console.Error.WriteLine("Checked " + engine.Name + ".");
        }

        IReadOnlyList<YamlIndependentReaders.Outcome> readers = YamlIndependentReaders.Run(suite, corpus, Option(args, "--python"));
        Console.Error.WriteLine("Checked independent readers.");
        string report = Write(suite, corpus, runs, details) + WriteReaders(readers, details);
        Console.WriteLine(report);
        if (outputPath != null) {
            File.WriteAllText(outputPath, report, new UTF8Encoding(false));
            Console.Error.WriteLine("Report written to " + Path.GetFullPath(outputPath) + ".");
        }

        return 0;
    }

    private static string Write(YamlTestSuite suite, IReadOnlyList<YamlCorpusFile> corpus, IReadOnlyList<EngineRun> runs, bool details) {
        var text = new StringBuilder();
        int errorCases = suite.Cases.Count(c => c.IsError);
        int jsonCases = suite.Cases.Count(c => !c.IsError && c.Json != null);
        text.AppendLine("# YAML conformance comparison");
        text.AppendLine();
        text.AppendLine("- yaml-test-suite " + suite.Commit.Substring(0, 12) + ": " + suite.Cases.Count + " cases (" + (suite.Cases.Count - errorCases) + " valid, " + errorCases + " error, " + jsonCases + " with JSON)");
        text.AppendLine("- Corpus: " + corpus.Count + " files, " + corpus.Sum(f => (long)f.Bytes.Length).ToString("N0", CultureInfo.InvariantCulture) + " bytes");
        text.AppendLine("- YamlDotNet " + Version(typeof(YamlDotNet.Core.Parser)) + ", SharpYaml " + Version(typeof(SharpYaml.Parser)) + ", VYaml " + VYamlConformanceEngine.PackageVersion + ", " + RuntimeInformation.FrameworkDescription + ", " + RuntimeInformation.OSDescription);
        text.AppendLine("- Cells are pass/applicable; failures (F) are wrong output or accepted invalid input, errors (E) are exceptions or timeouts on input the check expects to work.");
        text.AppendLine();

        text.AppendLine("## yaml-test-suite");
        text.AppendLine();
        text.AppendLine("| Engine | Events | Event content | Error cases | JSON | Round-trip bytes | Round-trip content | Emit |");
        text.AppendLine("|---|---|---|---|---|---|---|---|");
        foreach (EngineRun run in runs) {
            text.Append("| ").Append(run.Engine.Name).Append(Lossless(run.Engine));
            foreach (YamlConformanceCheck check in SuiteChecks) text.Append(" | ").Append(Cell(run.Suite, check));
            text.AppendLine(" |");
        }

        text.AppendLine();
        text.AppendLine("## Corpus round-trip");
        text.AppendLine();
        text.AppendLine("Each cell is accepts / bytes identical / content kept / emitted events re-parse.");
        text.AppendLine();
        text.Append("| File |");
        foreach (EngineRun run in runs) text.Append(' ').Append(run.Engine.Name).Append(" |");
        text.AppendLine();
        text.Append("|---|");
        foreach (EngineRun _ in runs) text.Append("---|");
        text.AppendLine();
        foreach (YamlCorpusFile file in corpus) {
            text.Append("| ").Append(file.Path).Append(" |");
            foreach (EngineRun run in runs) {
                IEnumerable<string> marks = CorpusChecks.Select(check => Mark(run.Corpus.First(r => r.Subject == file.Path && r.Check == check).Outcome));
                text.Append(' ').Append(string.Join(" / ", marks)).Append(" |");
            }

            text.AppendLine();
        }

        text.Append("| **Total** |");
        foreach (EngineRun run in runs) {
            IEnumerable<string> totals = CorpusChecks.Select(check => Cell(run.Corpus, check, compact: true));
            text.Append(' ').Append(string.Join(" / ", totals)).Append(" |");
        }

        text.AppendLine();
        if (details) AppendDetails(text, runs);
        return text.ToString();
    }

    private static string WriteReaders(IReadOnlyList<YamlIndependentReaders.Outcome> outcomes, bool details) {
        var text = new StringBuilder();
        text.AppendLine();
        text.AppendLine("## Independent readers on OfficeIMO.Yaml output");
        text.AppendLine();
        text.AppendLine("New documents are every valid suite case and corpus file emitted through `YamlWriter`; edited documents are the scripted-edit outputs. A reader counts for an input only when it reads the input as OfficeIMO.Yaml does.");
        text.AppendLine();
        text.AppendLine("| Reader | Output | Same data |");
        text.AppendLine("|---|---|---|");
        foreach (YamlIndependentReaders.Outcome outcome in outcomes) {
            string cell = outcome.Failures.Count > 0 && outcome.Failures[0].StartsWith("unavailable", StringComparison.Ordinal)
                ? outcome.Failures[0]
                : outcome.Agreed + "/" + outcome.Applicable;
            text.Append("| ").Append(outcome.Reader).Append(" | ").Append(outcome.Output).Append(" | ").Append(cell).AppendLine(" |");
        }
        if (details) {
            foreach (YamlIndependentReaders.Outcome outcome in outcomes) {
                foreach (string failure in outcome.Failures.Take(20)) text.Append("- ").Append(outcome.Reader).Append(' ').Append(outcome.Output).Append(": ").AppendLine(failure.Replace("\n", "\\n"));
            }
        }
        return text.ToString();
    }

    private static void AppendDetails(StringBuilder text, IReadOnlyList<EngineRun> runs) {
        foreach (EngineRun run in runs) {
            text.AppendLine();
            text.AppendLine("## Details: " + run.Engine.Name);
            foreach (YamlCheckResult result in run.Suite.Concat(run.Corpus)) {
                if (result.Outcome != YamlCheckOutcome.Fail && result.Outcome != YamlCheckOutcome.Error) continue;
                text.Append("- ").Append(result.Subject).Append(' ').Append(result.Check).Append(' ').Append(result.Outcome);
                if (result.Detail != null) text.Append(": ").Append(result.Detail.Replace("\r", "\\r").Replace("\n", "\\n"));
                text.AppendLine();
            }
        }
    }

    private static string Cell(IReadOnlyList<YamlCheckResult> results, YamlConformanceCheck check, bool compact = false) {
        int pass = 0, fail = 0, error = 0;
        foreach (YamlCheckResult result in results) {
            if (result.Check != check) continue;
            if (result.Outcome == YamlCheckOutcome.Pass) pass++;
            else if (result.Outcome == YamlCheckOutcome.Fail) fail++;
            else if (result.Outcome == YamlCheckOutcome.Error) error++;
        }

        int applicable = pass + fail + error;
        if (applicable == 0) return "n/a";
        string cell = pass + "/" + applicable;
        if (compact) return cell;
        var notes = new List<string>();
        if (fail > 0) notes.Add(fail + " F");
        if (error > 0) notes.Add(error + " E");
        return notes.Count == 0 ? cell : cell + " (" + string.Join(", ", notes) + ")";
    }

    private static string Mark(YamlCheckOutcome outcome) => outcome switch {
        YamlCheckOutcome.Pass => "yes",
        YamlCheckOutcome.Fail => "no",
        YamlCheckOutcome.Error => "err",
        _ => "-"
    };

    private static string Lossless(IYamlConformanceEngine engine) =>
        (engine.Features & YamlEngineFeatures.Lossless) != 0 ? " (claims lossless)" : string.Empty;

    private static string Version(Type type) =>
        type.Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion.Split('+')[0]
        ?? type.Assembly.GetName().Version?.ToString()
        ?? "unknown";

    private static string? Option(string[] args, string name) {
        for (int i = 0; i < args.Length - 1; i++) {
            if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase)) return args[i + 1];
        }

        return null;
    }

    private sealed class EngineRun {
        public EngineRun(IYamlConformanceEngine engine, IReadOnlyList<YamlCheckResult> suite, IReadOnlyList<YamlCheckResult> corpus) {
            Engine = engine;
            Suite = suite;
            Corpus = corpus;
        }

        public IYamlConformanceEngine Engine { get; }

        public IReadOnlyList<YamlCheckResult> Suite { get; }

        public IReadOnlyList<YamlCheckResult> Corpus { get; }
    }
}
