using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using OfficeIMO.Yaml.Benchmarks.Adapters;
using OfficeIMO.Yaml.Tests.Conformance;

namespace OfficeIMO.Yaml.Benchmarks.LargeCorpus;

/// <summary>
/// Runs OfficeIMO.Yaml and the reference libraries over every file of the large corpus:
/// <c>--large-corpus [--cache dir] [--output-dir dir] [--source id,id] [--limit n] [--timeout seconds]</c>.
/// Writes <c>large-corpus.md</c> (counts per source, category and check) and <c>large-corpus-failures.tsv</c>
/// (one line per failure with file and reason) to the output folder, by default <c>reports</c> in the cache.
/// </summary>
internal static class YamlLargeCorpusRun {
    private static readonly YamlLoadOptions Yaml11 = new() { Schema = YamlSchema.Yaml11 };
    private static readonly YamlReaderOptions StrictReader = new() { Strict = true };

    private enum Engine { OfficeIMO, YamlDotNet, SharpYaml, SharpYamlSyntax, VYaml }

    private enum Operation { Parse, Load, RoundTrip }

    public static int Run(string[] args) {
        string root = YamlLargeCorpus.ResolveRoot(Option(args, "--cache"));
        string outputDirectory = Path.GetFullPath(Option(args, "--output-dir") ?? Path.Combine(root, "reports"));
        TimeSpan timeout = TimeSpan.FromSeconds(double.Parse(Option(args, "--timeout") ?? "30", CultureInfo.InvariantCulture));
        string[]? sourceFilter = Option(args, "--source")?.Split(',', StringSplitOptions.RemoveEmptyEntries);
        int limit = int.Parse(Option(args, "--limit") ?? int.MaxValue.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);

        YamlLargeCorpus corpus = YamlLargeCorpus.Load(root);
        List<YamlLargeCorpusFile> files = corpus.Files.Where(f => sourceFilter == null || sourceFilter.Contains(f.Source)).Take(limit).ToList();
        var engines = new Dictionary<Engine, IYamlConformanceEngine> {
            [Engine.OfficeIMO] = new OfficeIMOYamlConformanceEngine(),
            [Engine.YamlDotNet] = new YamlDotNetConformanceEngine(),
            [Engine.SharpYaml] = new SharpYamlConformanceEngine(),
            [Engine.SharpYamlSyntax] = new SharpYamlSyntaxConformanceEngine(),
            [Engine.VYaml] = new VYamlConformanceEngine()
        };

        var run = new RunState(timeout);
        var watch = Stopwatch.StartNew();
        for (int i = 0; i < files.Count; i++) {
            run.Results.Add(CheckFile(files[i], engines, run));
            if ((i + 1) % 500 == 0) Console.Error.WriteLine("Checked " + (i + 1) + " of " + files.Count + " files (" + watch.Elapsed.TotalSeconds.ToString("F0", CultureInfo.InvariantCulture) + " s).");
        }

        Directory.CreateDirectory(outputDirectory);
        string reportPath = Path.Combine(outputDirectory, "large-corpus.md");
        string failuresPath = Path.Combine(outputDirectory, "large-corpus-failures.tsv");
        File.WriteAllText(reportPath, WriteReport(corpus, files, run, sourceFilter != null || limit != int.MaxValue), new UTF8Encoding(false));
        WriteFailures(failuresPath, run.Results);
        Console.WriteLine(File.ReadAllText(reportPath));
        Console.Error.WriteLine("Report: " + reportPath);
        Console.Error.WriteLine("Failures: " + failuresPath);
        return 0;
    }

    private static FileResult CheckFile(YamlLargeCorpusFile file, IReadOnlyDictionary<Engine, IYamlConformanceEngine> engines, RunState run) {
        var result = new FileResult(file);
        // A file that no longer matches the index stops the run: the results would not describe the pinned corpus.
        byte[] yaml = file.ReadYaml();
        result.YamlBytes = yaml.Length;
        foreach (Engine engine in new[] { Engine.OfficeIMO, Engine.YamlDotNet, Engine.SharpYaml, Engine.VYaml }) {
            IYamlConformanceEngine adapter = engines[engine];
            result.Events[engine] = run.Measure(engine, Operation.Parse, yaml.Length, () => adapter.Parse(yaml));
            result.Data[engine] = run.Measure(engine, Operation.Load, yaml.Length, () => adapter.Load(yaml));
        }

        foreach (Engine engine in new[] { Engine.OfficeIMO, Engine.YamlDotNet, Engine.SharpYaml, Engine.SharpYamlSyntax }) {
            IYamlConformanceEngine adapter = engines[engine];
            Outcome<byte[]> saved = run.Measure(engine, Operation.RoundTrip, yaml.Length, () => adapter.RoundTrip(yaml));
            result.RoundTrip[engine] = saved.Value == null ? null : yaml.AsSpan().SequenceEqual(saved.Value);
            if (engine == Engine.OfficeIMO && result.Events[Engine.OfficeIMO].Value != null) {
                if (saved.Value == null) result.Failures.Add(new Failure("round-trip", "OfficeIMO.Yaml", "error", saved.Error));
                else if (result.RoundTrip[engine] == false) result.Failures.Add(new Failure("round-trip", "OfficeIMO.Yaml", "fail", "first difference at byte " + FirstDifference(yaml, saved.Value) + " of " + yaml.Length));
            }
        }

        ClassifyAcceptance(result);
        CompareWithReferences(result);
        CompareSchemas(result, yaml, run);
        CompareStrict(result, yaml, run);
        return result;
    }

    private static void ClassifyAcceptance(FileResult result) {
        bool ours = result.Events[Engine.OfficeIMO].Value != null;
        bool yamlDotNet = result.Events[Engine.YamlDotNet].Value != null;
        bool sharpYaml = result.Events[Engine.SharpYaml].Value != null;
        string? ourError = result.Events[Engine.OfficeIMO].Error;
        if (ours) {
            result.Acceptance = yamlDotNet && sharpYaml ? Acceptance.AcceptedByAll
                : yamlDotNet || sharpYaml ? Acceptance.AcceptedSomeReferenceRejects
                : Acceptance.AcceptedOnlyByOfficeIMO;
        } else {
            result.Acceptance = yamlDotNet || sharpYaml ? Acceptance.RejectedOnlyByOfficeIMO : Acceptance.RejectedByAll;
        }

        switch (result.Acceptance) {
            case Acceptance.RejectedOnlyByOfficeIMO:
                result.Failures.Add(new Failure("parse", "OfficeIMO.Yaml", "investigate", ourError + " (accepted by " + string.Join(" and ", new[] { yamlDotNet ? "YamlDotNet" : null, sharpYaml ? "SharpYaml" : null }.Where(n => n != null)) + ")"));
                break;
            case Acceptance.AcceptedOnlyByOfficeIMO:
                result.Failures.Add(new Failure("parse", "OfficeIMO.Yaml", "investigate", "accepted while YamlDotNet (" + result.Events[Engine.YamlDotNet].Error + ") and SharpYaml (" + result.Events[Engine.SharpYaml].Error + ") reject"));
                break;
            case Acceptance.RejectedByAll:
                result.Failures.Add(new Failure("parse", "OfficeIMO.Yaml", "invalid", ourError));
                break;
        }

        if (ours && result.Data[Engine.OfficeIMO].Value == null) {
            bool referenceLoads = result.Data[Engine.YamlDotNet].Value != null || result.Data[Engine.SharpYaml].Value != null;
            result.Failures.Add(new Failure("load", "OfficeIMO.Yaml", referenceLoads ? "investigate" : "rejected", result.Data[Engine.OfficeIMO].Error));
        }
    }

    private static void CompareWithReferences(FileResult result) {
        IReadOnlyList<YamlTestEvent>? ourEvents = result.Events[Engine.OfficeIMO].Value;
        IReadOnlyList<YamlConformanceNode>? ourData = result.Data[Engine.OfficeIMO].Value;
        foreach (Engine engine in new[] { Engine.YamlDotNet, Engine.SharpYaml, Engine.VYaml }) {
            IReadOnlyList<YamlTestEvent>? theirEvents = result.Events[engine].Value;
            if (ourEvents != null && theirEvents != null) {
                int index = YamlTestEventText.FindContentDifference(ourEvents, theirEvents);
                result.EventsAgree[engine] = index < 0;
                if (index >= 0) {
                    string ours = index < ourEvents.Count ? ourEvents[index].ToString() : "<end>";
                    string theirs = index < theirEvents.Count ? theirEvents[index].ToString() : "<end>";
                    result.Failures.Add(new Failure("events", EngineName(engine), "differs", "event " + index + ": OfficeIMO '" + Clip(ours) + "' vs '" + Clip(theirs) + "'"));
                }
            }

            IReadOnlyList<YamlConformanceNode>? theirData = result.Data[engine].Value;
            if (ourData != null && theirData != null) {
                YamlDataDifference? difference = YamlDataComparison.First(ourData, theirData);
                result.DataDifference[engine] = difference;
                if (difference != null) result.Failures.Add(new Failure("data", EngineName(engine), difference.Kind.ToString(), difference.Describe()));
            }
        }
    }

    // How the same file loads under the YAML 1.1 schema, which several ecosystems (Ansible, older Kubernetes tooling,
    // PyYAML users) still apply without declaring it.
    private static void CompareSchemas(FileResult result, byte[] yaml, RunState run) {
        IReadOnlyList<YamlConformanceNode>? core = result.Data[Engine.OfficeIMO].Value;
        if (core == null) return;
        Outcome<IReadOnlyList<YamlConformanceNode>> yaml11 = run.Invoke(() => {
            YamlDocument document = YamlDocument.Load(new MemoryStream(yaml, writable: false), Yaml11);
            return (IReadOnlyList<YamlConformanceNode>)document.Documents.Select(part => OfficeIMOYamlConformanceEngine.ToConformance(part.Root)).ToList();
        });
        if (yaml11.Value == null) return;
        result.SchemaDifferences = YamlDataComparison.All(core, yaml11.Value, 1000);
    }

    // What the default mode accepts beyond the specification: files only it parses, its warnings and notes, and whether
    // each merge key it resolves without a directive names a mapping or a list of mappings.
    private static void CompareStrict(FileResult result, byte[] yaml, RunState run) {
        if (result.Events[Engine.OfficeIMO].Value == null) return;
        Outcome<object> strict = run.Invoke<object>(() => {
            using var reader = new YamlReader(new MemoryStream(yaml, writable: false), StrictReader);
            while (reader.Read()) { }
            return reader;
        });
        result.StrictRejects = strict.Value == null;
        if (result.Data[Engine.OfficeIMO].Value == null) return;
        run.Invoke<object>(() => {
            YamlDocument document = YamlDocument.Load(new MemoryStream(yaml, writable: false));
            result.LenientWarnings = document.Diagnostics.Count(d => d.Code == OfficeDiagnosticCodes.YamlLenientIndentation);
            result.MergeNotes = document.Diagnostics.Count(d => d.Code == OfficeDiagnosticCodes.YamlMergeKeyWithoutDirective);
            result.DuplicateKeyWarnings = document.Diagnostics.Count(d => d.Code == OfficeDiagnosticCodes.YamlDuplicateKey);
            if (result.MergeNotes == 0) return document;
            foreach (YamlDocumentPart part in document.Documents) {
                if (part.Schema != YamlSchema.Core) continue;
                foreach (YamlNode node in part.Root.Descendants()) {
                    if (node is not YamlMapping mapping) continue;
                    foreach (YamlMappingEntry entry in mapping) {
                        if (entry.Key is not YamlScalar { Kind: YamlScalarKind.Merge }) continue;
                        if (Merges(entry.Value)) result.MergeKeysResolved++;
                        else result.MergeKeysUnresolved++;
                    }
                }
            }
            return document;
        });
    }

    private static bool Merges(YamlNode value) => Target(value) switch {
        YamlMapping => true,
        YamlSequence list => list.All(item => Target(item) is YamlMapping),
        _ => false
    };

    private static YamlNode Target(YamlNode node) => node is YamlAlias alias ? alias.Target : node;

    // Labels a YAML 1.2 core versus 1.1 scalar difference by the 1.1 rule that caused it.
    private static string SchemaRule(YamlDataDifference difference) {
        YamlConformanceNode? core = difference.Left;
        YamlConformanceNode? yaml11 = difference.Right;
        if (core == null || yaml11 == null) return "other";
        string text = core.Scalar ?? string.Empty;
        if (core.Kind == YamlConformanceNodeKind.String && yaml11.Kind == YamlConformanceNodeKind.Boolean) return "y/n/yes/no/on/off booleans";
        if (core.Kind == YamlConformanceNodeKind.String && yaml11.Kind is YamlConformanceNodeKind.Integer or YamlConformanceNodeKind.Float) {
            if (text.IndexOf(':') >= 0) return "base-60 numbers (1:30)";
            if (text.IndexOf('_') >= 0) return "underscore digit separators";
            if (text.StartsWith("0b", StringComparison.Ordinal) || text.StartsWith("+0b", StringComparison.Ordinal) || text.StartsWith("-0b", StringComparison.Ordinal)) return "0b binary integers";
            return "other numbers";
        }

        if (core.Kind == YamlConformanceNodeKind.Integer && yaml11.Kind == YamlConformanceNodeKind.Integer) return "leading-zero octal (0755)";
        if (core.Kind == YamlConformanceNodeKind.Integer && yaml11.Kind == YamlConformanceNodeKind.String) return "0o octal is a string in 1.1";
        if (core.Kind == YamlConformanceNodeKind.Float && yaml11.Kind == YamlConformanceNodeKind.String) return "floats 1.1 does not recognize (1e3, .5)";
        if (core.Kind == YamlConformanceNodeKind.String && yaml11.Kind == YamlConformanceNodeKind.Null) return "null spellings";
        return core.Kind + " to " + yaml11.Kind;
    }

    private static string WriteReport(YamlLargeCorpus corpus, IReadOnlyList<YamlLargeCorpusFile> files, RunState run, bool partial) {
        IReadOnlyList<FileResult> results = run.Results;
        var text = new StringBuilder();
        text.AppendLine("# YAML large-corpus run");
        text.AppendLine();
        text.AppendLine("- " + files.Count + " files" + (partial ? " (filtered run)" : string.Empty) + " from " + files.Select(f => f.Source).Distinct().Count() + " sources, " + YamlLargeCorpus.Format(results.Sum(r => r.YamlBytes)) + " bytes of YAML; " + corpus.Exclusions.Count + " indexed files excluded before the run");
        text.AppendLine("- YamlDotNet " + Version(typeof(YamlDotNet.Core.Parser)) + ", SharpYaml " + Version(typeof(SharpYaml.Parser)) + ", VYaml " + VYamlConformanceEngine.PackageVersion + ", " + RuntimeInformation.FrameworkDescription + ", " + RuntimeInformation.OSDescription + ", per-call timeout " + run.Timeout.TotalSeconds.ToString(CultureInfo.InvariantCulture) + " s");
        text.AppendLine("- Acceptance is classified by YamlDotNet and SharpYaml. VYaml events carry no presentation, so all event comparisons are by content: kinds, values, tags, anchors and aliases. Data compares the loaded YAML 1.2 core-schema values; VYaml resolves scalars itself.");
        text.AppendLine();

        text.AppendLine("## Checks");
        text.AppendLine();
        text.AppendLine("Cells are passes/applicable. Events and data compare each reference with OfficeIMO.Yaml on files both accept.");
        text.AppendLine();
        text.AppendLine("| Check | OfficeIMO.Yaml | YamlDotNet | SharpYaml | SharpYaml syntax | VYaml |");
        text.AppendLine("|---|---|---|---|---|---|");
        text.AppendLine("| Parses | " + Ratio(results, r => r.Events[Engine.OfficeIMO].Value != null) + " | " + Ratio(results, r => r.Events[Engine.YamlDotNet].Value != null) + " | " + Ratio(results, r => r.Events[Engine.SharpYaml].Value != null) + " | n/a | " + Ratio(results, r => r.Events[Engine.VYaml].Value != null) + " |");
        text.AppendLine("| Loads | " + Ratio(results, r => r.Data[Engine.OfficeIMO].Value != null) + " | " + Ratio(results, r => r.Data[Engine.YamlDotNet].Value != null) + " | " + Ratio(results, r => r.Data[Engine.SharpYaml].Value != null) + " | n/a | " + Ratio(results, r => r.Data[Engine.VYaml].Value != null) + " |");
        text.AppendLine("| Round-trip bytes | " + RoundTripCell(results, Engine.OfficeIMO) + " | " + RoundTripCell(results, Engine.YamlDotNet) + " | " + RoundTripCell(results, Engine.SharpYaml) + " | " + RoundTripCell(results, Engine.SharpYamlSyntax) + " | n/a |");
        text.AppendLine("| Events agree with OfficeIMO.Yaml | - | " + AgreeCell(results, Engine.YamlDotNet, events: true) + " | " + AgreeCell(results, Engine.SharpYaml, events: true) + " | n/a | " + AgreeCell(results, Engine.VYaml, events: true) + " |");
        text.AppendLine("| Data agrees with OfficeIMO.Yaml | - | " + AgreeCell(results, Engine.YamlDotNet, events: false) + " | " + AgreeCell(results, Engine.SharpYaml, events: false) + " | n/a | " + AgreeCell(results, Engine.VYaml, events: false) + " |");
        text.AppendLine();

        text.AppendLine("## Acceptance");
        text.AppendLine();
        text.AppendLine("| Class | Files | Meaning |");
        text.AppendLine("|---|---|---|");
        foreach (Acceptance acceptance in Enum.GetValues<Acceptance>()) {
            text.AppendLine("| " + acceptance + " | " + results.Count(r => r.Acceptance == acceptance) + " | " + Meaning(acceptance) + " |");
        }

        text.AppendLine();
        AppendErrorCodes(text, results);
        AppendGroups(text, "Sources", "Source", results, r => r.File.Source, corpus);
        AppendGroups(text, "Categories", "Category", results, r => r.File.Category, null);
        AppendDifferences(text, results);
        AppendSchemaDifferences(text, results);
        AppendStrict(text, results);
        AppendExclusions(text, corpus, files);
        AppendThroughput(text, run);
        return text.ToString();
    }

    private static void AppendErrorCodes(StringBuilder text, IReadOnlyList<FileResult> results) {
        var rejected = results.Where(r => r.Acceptance is Acceptance.RejectedByAll or Acceptance.RejectedOnlyByOfficeIMO).ToList();
        if (rejected.Count == 0) return;
        text.AppendLine("OfficeIMO.Yaml errors on rejected files:");
        text.AppendLine();
        text.AppendLine("| Error | Rejected by all | Rejected only by OfficeIMO.Yaml |");
        text.AppendLine("|---|---|---|");
        foreach (var group in rejected.GroupBy(r => ErrorCode(r.Events[Engine.OfficeIMO].Error)).OrderByDescending(g => g.Count())) {
            text.AppendLine("| " + group.Key + " | " + group.Count(r => r.Acceptance == Acceptance.RejectedByAll) + " | " + group.Count(r => r.Acceptance == Acceptance.RejectedOnlyByOfficeIMO) + " |");
        }

        text.AppendLine();
    }

    private static void AppendGroups(StringBuilder text, string title, string column, IReadOnlyList<FileResult> results, Func<FileResult, string> key, YamlLargeCorpus? corpus) {
        text.AppendLine("## " + title);
        text.AppendLine();
        text.AppendLine("| " + column + (corpus != null ? " | License" : string.Empty) + " | Files | Bytes | Parses | Loads | Round-trip exact | Invalid (all reject) | Investigate | YamlDotNet data differs | VYaml data differs |");
        text.AppendLine("|---|" + (corpus != null ? "---|" : string.Empty) + "---|---|---|---|---|---|---|---|---|");
        foreach (var group in results.GroupBy(key).OrderBy(g => g.Key, StringComparer.Ordinal)) {
            text.Append("| ").Append(group.Key);
            if (corpus != null) text.Append(" | ").Append(corpus.Sources.FirstOrDefault(s => s.Id == group.Key)?.License ?? "?");
            text.Append(" | ").Append(group.Count());
            text.Append(" | ").Append(YamlLargeCorpus.Format(group.Sum(r => r.YamlBytes)));
            text.Append(" | ").Append(group.Count(r => r.Events[Engine.OfficeIMO].Value != null));
            text.Append(" | ").Append(group.Count(r => r.Data[Engine.OfficeIMO].Value != null));
            text.Append(" | ").Append(group.Count(r => r.RoundTrip.TryGetValue(Engine.OfficeIMO, out bool? exact) && exact == true));
            text.Append(" | ").Append(group.Count(r => r.Acceptance == Acceptance.RejectedByAll));
            text.Append(" | ").Append(group.Count(r => r.Failures.Any(f => f.Outcome == "investigate")));
            text.Append(" | ").Append(group.Count(r => r.DataDifference.TryGetValue(Engine.YamlDotNet, out YamlDataDifference? d) && d != null));
            text.Append(" | ").Append(group.Count(r => r.DataDifference.TryGetValue(Engine.VYaml, out YamlDataDifference? d) && d != null));
            text.AppendLine(" |");
        }

        text.AppendLine();
    }

    private static void AppendDifferences(StringBuilder text, IReadOnlyList<FileResult> results) {
        text.AppendLine("## Differences from OfficeIMO.Yaml");
        text.AppendLine();
        text.AppendLine("First difference per file, grouped by kind and by the kinds on each side. Samples are OfficeIMO.Yaml value vs reference value.");
        text.AppendLine();
        text.AppendLine("| Engine | Difference | Files | Sample |");
        text.AppendLine("|---|---|---|---|");
        foreach (Engine engine in new[] { Engine.YamlDotNet, Engine.SharpYaml, Engine.VYaml }) {
            int eventDiffs = results.Count(r => r.EventsAgree.TryGetValue(engine, out bool agree) && !agree);
            if (eventDiffs > 0) {
                FileResult sample = results.First(r => r.EventsAgree.TryGetValue(engine, out bool agree) && !agree);
                text.AppendLine("| " + EngineName(engine) + " | events | " + eventDiffs + " | " + Cell(sample.File.Name + " " + sample.Failures.First(f => f.Check == "events" && f.Engine == EngineName(engine)).Detail) + " |");
            }

            var differences = results.Where(r => r.DataDifference.TryGetValue(engine, out YamlDataDifference? d) && d != null).Select(r => (File: r.File, Difference: r.DataDifference[engine]!));
            foreach (var group in differences.GroupBy(d => d.Difference.Kind + ": " + d.Difference.Left?.Kind + " vs " + d.Difference.Right?.Kind).OrderByDescending(g => g.Count())) {
                var first = group.First();
                text.AppendLine("| " + EngineName(engine) + " | " + group.Key + " | " + group.Count() + " | " + Cell(first.File.Name + " " + first.Difference.Path + ": " + YamlDataDifference.Show(first.Difference.Left) + " vs " + YamlDataDifference.Show(first.Difference.Right)) + " |");
            }
        }

        text.AppendLine();
    }

    private static void AppendSchemaDifferences(StringBuilder text, IReadOnlyList<FileResult> results) {
        var affected = results.Where(r => r.SchemaDifferences.Count > 0).ToList();
        text.AppendLine("## YAML 1.2 core versus 1.1 schema");
        text.AppendLine();
        text.AppendLine(affected.Count + " of " + results.Count(r => r.Data[Engine.OfficeIMO].Value != null) + " loaded files hold values that a YAML 1.1 loader resolves differently. The YamlDotNet, SharpYaml and VYaml columns above all use 1.2 core rules, so these are not counted as differences there.");
        text.AppendLine();
        text.AppendLine("| 1.1 rule | Files | Values | Sample (1.2 core vs 1.1) |");
        text.AppendLine("|---|---|---|---|");
        var all = affected.SelectMany(r => r.SchemaDifferences.Select(d => (File: r.File, Difference: d)));
        foreach (var group in all.GroupBy(d => SchemaRule(d.Difference)).OrderByDescending(g => g.Select(d => d.File).Distinct().Count())) {
            var first = group.First();
            text.AppendLine("| " + group.Key + " | " + group.Select(d => d.File).Distinct().Count() + " | " + group.Count() + " | " + Cell(first.File.Name + " " + first.Difference.Path + ": " + YamlDataDifference.Show(first.Difference.Left) + " vs " + YamlDataDifference.Show(first.Difference.Right)) + " |");
        }

        text.AppendLine();
        text.AppendLine("| Category | Files affected by 1.1 rules |");
        text.AppendLine("|---|---|");
        foreach (var group in affected.GroupBy(r => r.File.Category).OrderByDescending(g => g.Count())) text.AppendLine("| " + group.Key + " | " + group.Count() + " |");
        text.AppendLine();
    }

    private static void AppendStrict(StringBuilder text, IReadOnlyList<FileResult> results) {
        var onlyDefault = results.Where(r => r.StrictRejects).ToList();
        var merged = results.Where(r => r.MergeNotes > 0).ToList();
        text.AppendLine("## Default mode and Strict");
        text.AppendLine();
        text.AppendLine("The runs above use the default mode. Strict follows the YAML 1.2.2 specification exactly.");
        text.AppendLine();
        text.AppendLine("| Measure | Count |");
        text.AppendLine("|---|---|");
        text.AppendLine("| Files Strict rejects and the default mode parses | " + onlyDefault.Count + " |");
        text.AppendLine("| of those, loaded data agrees with YamlDotNet | " + onlyDefault.Count(r => r.DataDifference.TryGetValue(Engine.YamlDotNet, out YamlDataDifference? d) && d == null) + "/" + onlyDefault.Count(r => r.DataDifference.ContainsKey(Engine.YamlDotNet)) + " |");
        text.AppendLine("| of those, loaded data agrees with SharpYaml | " + onlyDefault.Count(r => r.DataDifference.TryGetValue(Engine.SharpYaml, out YamlDataDifference? d) && d == null) + "/" + onlyDefault.Count(r => r.DataDifference.ContainsKey(Engine.SharpYaml)) + " |");
        text.AppendLine("| of those, round-trip exact | " + onlyDefault.Count(r => r.RoundTrip.TryGetValue(Engine.OfficeIMO, out bool? exact) && exact == true) + "/" + onlyDefault.Count + " |");
        text.AppendLine("| Files with yaml-lenient-indentation warnings (warnings) | " + results.Count(r => r.LenientWarnings > 0) + " (" + results.Sum(r => r.LenientWarnings) + ") |");
        text.AppendLine("| Files with yaml-duplicate-key warnings (warnings) | " + results.Count(r => r.DuplicateKeyWarnings > 0) + " (" + results.Sum(r => r.DuplicateKeyWarnings) + ") |");
        text.AppendLine("| Files with yaml-merge-key-without-directive notes (notes) | " + merged.Count + " (" + merged.Sum(r => r.MergeNotes) + ") |");
        text.AppendLine("| Merge keys resolving to a mapping or list of mappings | " + merged.Sum(r => r.MergeKeysResolved) + "/" + merged.Sum(r => r.MergeKeysResolved + r.MergeKeysUnresolved) + " |");
        text.AppendLine();
        if (onlyDefault.Count > 0) {
            text.AppendLine("| Source | Files Strict rejects | Merge-key files |");
            text.AppendLine("|---|---|---|");
            foreach (var group in results.Where(r => r.StrictRejects || r.MergeNotes > 0).GroupBy(r => r.File.Source).OrderBy(g => g.Key, StringComparer.Ordinal)) {
                text.AppendLine("| " + group.Key + " | " + group.Count(r => r.StrictRejects) + " | " + group.Count(r => r.MergeNotes > 0) + " |");
            }

            text.AppendLine();
        }
    }

    private static void AppendExclusions(StringBuilder text, YamlLargeCorpus corpus, IReadOnlyList<YamlLargeCorpusFile> files) {
        var sources = new HashSet<string>(files.Select(f => f.Source), StringComparer.Ordinal);
        var exclusions = corpus.Exclusions.Where(e => sources.Contains(e.Source)).ToList();
        if (exclusions.Count == 0) return;
        text.AppendLine("## Excluded before the run");
        text.AppendLine();
        text.AppendLine("| Source | Reason | Files |");
        text.AppendLine("|---|---|---|");
        foreach (var group in exclusions.GroupBy(e => (e.Source, e.Reason)).OrderBy(g => g.Key.Source, StringComparer.Ordinal)) text.AppendLine("| " + group.Key.Source + " | " + group.Key.Reason + " | " + group.Count() + " |");
        text.AppendLine();
    }

    private static void AppendThroughput(StringBuilder text, RunState run) {
        text.AppendLine("## Throughput");
        text.AppendLine();
        text.AppendLine("Single-threaded totals over the files each library completed, through the conformance adapters (so they include converting to the shared event and value form). Use the BenchmarkDotNet lanes for comparable measurements.");
        text.AppendLine();
        text.AppendLine("| Library | Operation | Files | MB | Seconds | MB/s |");
        text.AppendLine("|---|---|---|---|---|---|");
        foreach (KeyValuePair<(Engine Engine, Operation Operation), Totals> entry in run.Totals.OrderBy(e => e.Key.Operation).ThenBy(e => e.Key.Engine)) {
            double megabytes = entry.Value.Bytes / 1048576.0;
            double seconds = entry.Value.Ticks / (double)Stopwatch.Frequency;
            text.AppendLine("| " + EngineName(entry.Key.Engine) + " | " + entry.Key.Operation + " | " + entry.Value.Files + " | " + megabytes.ToString("F1", CultureInfo.InvariantCulture) + " | " + seconds.ToString("F2", CultureInfo.InvariantCulture) + " | " + (seconds > 0 ? (megabytes / seconds).ToString("F1", CultureInfo.InvariantCulture) : "-") + " |");
        }

        text.AppendLine();
    }

    private static void WriteFailures(string path, IReadOnlyList<FileResult> results) {
        using var writer = new StreamWriter(path, false, new UTF8Encoding(false));
        writer.WriteLine("source\tcategory\tpath\tcheck\tengine\toutcome\tdetail");
        foreach (FileResult result in results) {
            foreach (Failure failure in result.Failures) {
                writer.WriteLine(string.Join("\t", result.File.Source, result.File.Category, result.File.Path, failure.Check, failure.Engine, failure.Outcome, Tsv(failure.Detail)));
            }
        }
    }

    private static string Ratio(IReadOnlyList<FileResult> results, Func<FileResult, bool> predicate) => results.Count(predicate) + "/" + results.Count;

    private static string RoundTripCell(IReadOnlyList<FileResult> results, Engine engine) {
        int applicable = results.Count(r => r.Events[Engine.OfficeIMO].Value != null || r.Events[Engine.YamlDotNet].Value != null || r.Events[Engine.SharpYaml].Value != null);
        int exact = results.Count(r => r.RoundTrip.TryGetValue(engine, out bool? same) && same == true);
        int errors = results.Count(r => (r.Events[Engine.OfficeIMO].Value != null || r.Events[Engine.YamlDotNet].Value != null || r.Events[Engine.SharpYaml].Value != null) && r.RoundTrip.TryGetValue(engine, out bool? same) && same == null);
        return exact + "/" + applicable + (errors > 0 ? " (" + errors + " E)" : string.Empty);
    }

    private static string AgreeCell(IReadOnlyList<FileResult> results, Engine engine, bool events) {
        if (events) {
            var applicable = results.Where(r => r.EventsAgree.ContainsKey(engine)).ToList();
            return applicable.Count(r => r.EventsAgree[engine]) + "/" + applicable.Count;
        }

        var loaded = results.Where(r => r.DataDifference.ContainsKey(engine)).ToList();
        return loaded.Count(r => r.DataDifference[engine] == null) + "/" + loaded.Count;
    }

    private static string Meaning(Acceptance acceptance) => acceptance switch {
        Acceptance.AcceptedByAll => "OfficeIMO.Yaml, YamlDotNet and SharpYaml parse it",
        Acceptance.AcceptedSomeReferenceRejects => "OfficeIMO.Yaml and one reference parse it",
        Acceptance.AcceptedOnlyByOfficeIMO => "both references reject it; investigate whether OfficeIMO.Yaml is too lenient",
        Acceptance.RejectedByAll => "invalid YAML confirmed by both references",
        Acceptance.RejectedOnlyByOfficeIMO => "a reference parses it; investigate as a potential OfficeIMO.Yaml bug",
        _ => string.Empty
    };

    private static string ErrorCode(string? error) {
        if (error == null) return "unknown";
        int bracket = error.IndexOf('[');
        int close = bracket < 0 ? -1 : error.IndexOf(']', bracket);
        return bracket >= 0 && close > bracket ? error.Substring(bracket + 1, close - bracket - 1) : error.Split(':')[0];
    }

    private static string EngineName(Engine engine) => engine switch {
        Engine.OfficeIMO => "OfficeIMO.Yaml",
        Engine.SharpYamlSyntax => "SharpYaml syntax",
        _ => engine.ToString()
    };

    private static string Describe(Exception exception) {
        Exception inner = exception is AggregateException aggregate && aggregate.InnerExceptions.Count == 1 ? aggregate.InnerExceptions[0] : exception;
        string message = inner.Message;
        int newline = message.IndexOf('\n');
        if (newline >= 0) message = message.Substring(0, newline).TrimEnd('\r');
        if (message.Length > 200) message = message.Substring(0, 200) + "...";
        string code = inner is YamlException yaml ? " [" + yaml.Code + "]" : string.Empty;
        return inner.GetType().Name + code + ": " + message;
    }

    private static string Clip(string text) => text.Length > 80 ? text.Substring(0, 77) + "..." : text;

    private static string Cell(string text) => Clip(text).Replace("|", "\\|").Replace("\n", "\\n");

    private static string Tsv(string? text) => (text ?? string.Empty).Replace("\t", "\\t").Replace("\r", "\\r").Replace("\n", "\\n");

    private static int FirstDifference(byte[] expected, byte[] actual) {
        int count = Math.Min(expected.Length, actual.Length);
        for (int i = 0; i < count; i++) {
            if (expected[i] != actual[i]) return i;
        }

        return count;
    }

    private static string Version(Type type) =>
        type.Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion.Split('+')[0] ?? type.Assembly.GetName().Version?.ToString() ?? "unknown";

    private static string? Option(string[] args, string name) {
        for (int i = 0; i < args.Length - 1; i++) {
            if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase)) return args[i + 1];
        }

        return null;
    }

    private enum Acceptance { AcceptedByAll, AcceptedSomeReferenceRejects, AcceptedOnlyByOfficeIMO, RejectedByAll, RejectedOnlyByOfficeIMO }

    private sealed record Failure(string Check, string Engine, string Outcome, string? Detail);

    private readonly struct Outcome<T> where T : class {
        public Outcome(T? value, string? error) {
            Value = value;
            Error = error;
        }

        public T? Value { get; }

        public string? Error { get; }
    }

    private sealed class Totals {
        public int Files;
        public long Bytes;
        public long Ticks;
    }

    private sealed class RunState {
        public RunState(TimeSpan timeout) => Timeout = timeout;

        public TimeSpan Timeout { get; }

        public List<FileResult> Results { get; } = new();

        public Dictionary<(Engine Engine, Operation Operation), Totals> Totals { get; } = new();

        /// <summary>Runs one call with the timeout and adds its elapsed time to the totals when it completes.</summary>
        public Outcome<T> Measure<T>(Engine engine, Operation operation, long bytes, Func<T> action) where T : class {
            long ticks = 0;
            Outcome<T> outcome = Invoke(() => {
                long start = Stopwatch.GetTimestamp();
                T value = action();
                ticks = Stopwatch.GetTimestamp() - start;
                return value;
            });
            if (outcome.Value != null) {
                if (!Totals.TryGetValue((engine, operation), out Totals? totals)) Totals[(engine, operation)] = totals = new Totals();
                totals.Files++;
                totals.Bytes += bytes;
                totals.Ticks += ticks;
            }

            return outcome;
        }

        public Outcome<T> Invoke<T>(Func<T> action) where T : class {
            try {
                Task<T> task = Task.Run(action);
                if (!task.Wait(Timeout)) return new Outcome<T>(null, "timed out after " + Timeout.TotalSeconds.ToString(CultureInfo.InvariantCulture) + " s");
                return new Outcome<T>(task.Result, null);
            } catch (Exception exception) {
                return new Outcome<T>(null, YamlLargeCorpusRun.Describe(exception));
            }
        }
    }

    private sealed class FileResult {
        public FileResult(YamlLargeCorpusFile file) => File = file;

        public YamlLargeCorpusFile File { get; }

        public long YamlBytes { get; set; }

        public Acceptance Acceptance { get; set; }

        public Dictionary<Engine, Outcome<IReadOnlyList<YamlTestEvent>>> Events { get; } = new();

        public Dictionary<Engine, Outcome<IReadOnlyList<YamlConformanceNode>>> Data { get; } = new();

        /// <summary>True when the saved bytes equal the input, false when they differ, null when saving failed.</summary>
        public Dictionary<Engine, bool?> RoundTrip { get; } = new();

        public Dictionary<Engine, bool> EventsAgree { get; } = new();

        public Dictionary<Engine, YamlDataDifference?> DataDifference { get; } = new();

        public IReadOnlyList<YamlDataDifference> SchemaDifferences { get; set; } = Array.Empty<YamlDataDifference>();

        public bool StrictRejects { get; set; }

        public int LenientWarnings { get; set; }

        public int MergeNotes { get; set; }

        public int DuplicateKeyWarnings { get; set; }

        public int MergeKeysResolved { get; set; }

        public int MergeKeysUnresolved { get; set; }

        public List<Failure> Failures { get; } = new();
    }
}
