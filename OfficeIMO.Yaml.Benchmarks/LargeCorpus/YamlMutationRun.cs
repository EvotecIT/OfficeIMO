using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using OfficeIMO.Yaml.Tests.Conformance;

namespace OfficeIMO.Yaml.Benchmarks.LargeCorpus;

/// <summary>
/// Hostile-input run over mutated large-corpus files:
/// <c>--mutations [--cache dir] [--output-dir dir] [--seed n] [--files n] [--per-kind n] [--timeout seconds] [--file source:path] [--kind name]</c>.
/// Each mutant goes through the forward-only reader and the document model (load, unchanged save, resolved view). A
/// mutant passes when only <see cref="YamlException"/> (including <see cref="YamlLimitExceededException"/>) or
/// <see cref="OperationCanceledException"/> escapes, nothing hangs past the timeout, and allocation stays bounded.
/// Mutants derive from the seed, file name, kind and variant number only, so a filtered run reproduces the same bytes;
/// failing mutants are also saved under <c>mutants</c> in the output folder.
/// </summary>
internal static class YamlMutationRun {
    private const long AllocationFloor = 512L * 1024 * 1024;
    private const int AllocationFactor = 100;

    internal enum MutationKind {
        Truncate, ByteFlip, IndicatorInsert, LineEdit, IndentShift, DeepFlowNesting, DeepBlockNesting,
        AliasBomb, AnchorChaos, HugeScalar, NodeFlood, InvalidUtf8, ControlCharacters, EncodingSwap
    }

    private enum Verdict { Accepted, Rejected, LimitExceeded, Canceled, Crash, Hang, Memory }

    public static int Run(string[] args) {
        string root = YamlLargeCorpus.ResolveRoot(Option(args, "--cache"));
        string outputDirectory = Path.GetFullPath(Option(args, "--output-dir") ?? Path.Combine(root, "reports"));
        int seed = int.Parse(Option(args, "--seed") ?? "20261110", CultureInfo.InvariantCulture);
        int fileCount = int.Parse(Option(args, "--files") ?? "300", CultureInfo.InvariantCulture);
        int perKind = int.Parse(Option(args, "--per-kind") ?? "2", CultureInfo.InvariantCulture);
        TimeSpan timeout = TimeSpan.FromSeconds(double.Parse(Option(args, "--timeout") ?? "10", CultureInfo.InvariantCulture));
        string? onlyFile = Option(args, "--file");
        MutationKind[] kinds = Option(args, "--kind") is string kindName ? new[] { Enum.Parse<MutationKind>(kindName, ignoreCase: true) } : Enum.GetValues<MutationKind>();

        YamlLargeCorpus corpus = YamlLargeCorpus.Load(root);
        List<YamlLargeCorpusFile> files = SelectFiles(corpus.Files, seed, fileCount, onlyFile);
        string mutantDirectory = Path.Combine(outputDirectory, "mutants");
        Directory.CreateDirectory(mutantDirectory);

        var results = new List<MutantResult>();
        var watch = Stopwatch.StartNew();
        foreach (YamlLargeCorpusFile file in files) {
            byte[] input = file.ReadYaml();
            foreach (MutationKind kind in kinds) {
                for (int variant = 0; variant < perKind; variant++) {
                    var random = new Random(StableSeed(seed, file.Name, kind, variant));
                    byte[] mutant = Mutate(input, kind, random, out string description);
                    string id = seed.ToString(CultureInfo.InvariantCulture) + "|" + file.Name + "|" + kind + "|" + variant.ToString(CultureInfo.InvariantCulture);
                    var result = new MutantResult(id, file, kind, description, mutant.Length,
                        Execute(() => ReadAll(mutant), timeout, mutant.Length),
                        Execute(() => LoadSaveResolve(mutant), timeout, mutant.Length));
                    if (result.Reader.IsFailure || result.Model.IsFailure || result.RoundTripMismatch) {
                        File.WriteAllBytes(Path.Combine(mutantDirectory, SafeName(id) + ".yaml"), mutant);
                    }

                    results.Add(result);
                }
            }

            if (results.Count % 1000 < kinds.Length * perKind) Console.Error.WriteLine("Ran " + results.Count + " mutants (" + watch.Elapsed.TotalSeconds.ToString("F0", CultureInfo.InvariantCulture) + " s).");
        }

        Directory.CreateDirectory(outputDirectory);
        string reportPath = Path.Combine(outputDirectory, "mutations.md");
        string failuresPath = Path.Combine(outputDirectory, "mutation-failures.tsv");
        File.WriteAllText(reportPath, WriteReport(results, seed, files.Count, perKind, timeout, watch.Elapsed), new UTF8Encoding(false));
        using (var writer = new StreamWriter(failuresPath, false, new UTF8Encoding(false))) {
            writer.WriteLine("id\tstage\tverdict\tdetail\tmutation");
            foreach (MutantResult result in results) {
                if (result.Reader.IsFailure) writer.WriteLine(result.Id + "\treader\t" + result.Reader.Verdict + "\t" + Tsv(result.Reader.Detail) + "\t" + Tsv(result.Description));
                if (result.Model.IsFailure) writer.WriteLine(result.Id + "\tmodel\t" + result.Model.Verdict + "\t" + Tsv(result.Model.Detail) + "\t" + Tsv(result.Description));
                if (result.RoundTripMismatch) writer.WriteLine(result.Id + "\tmodel\tRoundTripMismatch\t" + Tsv(result.Model.Detail) + "\t" + Tsv(result.Description));
            }
        }

        Console.WriteLine(File.ReadAllText(reportPath));
        Console.Error.WriteLine("Report: " + reportPath);
        return results.Any(r => r.Reader.IsFailure || r.Model.IsFailure) ? 1 : 0;
    }

    private static List<YamlLargeCorpusFile> SelectFiles(IReadOnlyList<YamlLargeCorpusFile> all, int seed, int count, string? onlyFile) {
        List<YamlLargeCorpusFile> sorted = all.OrderBy(f => f.Name, StringComparer.Ordinal).ToList();
        if (onlyFile != null) return sorted.Where(f => f.Name == onlyFile).ToList();
        var random = new Random(seed);
        for (int i = sorted.Count - 1; i > 0; i--) {
            int j = random.Next(i + 1);
            (sorted[i], sorted[j]) = (sorted[j], sorted[i]);
        }

        return sorted.Take(count).ToList();
    }

    private static void ReadAll(byte[] yaml) {
        using var reader = new YamlReader(new MemoryStream(yaml, writable: false), options: null, CurrentToken.Value);
        while (reader.Read()) {
            switch (reader.TokenType) {
                case YamlTokenType.Scalar:
                    _ = reader.GetString();
                    _ = reader.Anchor;
                    _ = reader.Tag;
                    break;
                case YamlTokenType.Alias:
                    _ = reader.GetString();
                    break;
                case YamlTokenType.MappingStart:
                case YamlTokenType.SequenceStart:
                    _ = reader.Anchor;
                    _ = reader.Tag;
                    break;
            }
        }
    }

    // Returns null when the unchanged save reproduces the input, or a description of the first difference.
    private static string? LoadSaveResolve(byte[] yaml) {
        YamlDocument document = YamlDocument.LoadAsync(new MemoryStream(yaml, writable: false), null, CurrentToken.Value).GetAwaiter().GetResult();
        var saved = new MemoryStream();
        document.Save(saved);
        foreach (YamlDocumentPart part in document.Documents) _ = OfficeIMOYamlConformanceEngine.ToConformance(part.Root);
        byte[] output = saved.ToArray();
        if (output.AsSpan().SequenceEqual(yaml)) return null;
        int count = Math.Min(output.Length, yaml.Length);
        int index = 0;
        while (index < count && output[index] == yaml[index]) index++;
        return "saved output differs at byte " + index + " of " + yaml.Length;
    }

    private static readonly AsyncLocal<CancellationToken> CurrentToken = new();

    // Sampled after every stage; Process.PeakWorkingSet64 is not reported on every platform.
    private static long PeakWorkingSet;

    // The stage's cancellation token, set on the worker thread so the reader and LoadAsync observe the timeout.
    private static StageResult Execute(Action action, TimeSpan timeout, int inputLength) =>
        Execute(() => {
            action();
            return null;
        }, timeout, inputLength);

    private static StageResult Execute(Func<string?> action, TimeSpan timeout, int inputLength) {
        using var cancellation = new CancellationTokenSource(timeout);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: false);
        var watch = Stopwatch.StartNew();
        Task<string?> task = Task.Run(() => {
            CurrentToken.Value = cancellation.Token;
            return action();
        });
        bool finished;
        try {
            finished = task.Wait(timeout + TimeSpan.FromSeconds(5));
        } catch (AggregateException) {
            finished = true;
        }

        watch.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: false) - allocatedBefore;
        PeakWorkingSet = Math.Max(PeakWorkingSet, Environment.WorkingSet);
        if (!finished) return new StageResult(Verdict.Hang, "still running " + (timeout.TotalSeconds + 5).ToString(CultureInfo.InvariantCulture) + " s after start", watch.Elapsed, allocated, null);
        if (task.IsFaulted) {
            Exception exception = task.Exception!.InnerExceptions.Count == 1 ? task.Exception.InnerExceptions[0] : task.Exception;
            Verdict verdict = exception switch {
                YamlLimitExceededException => Verdict.LimitExceeded,
                YamlException => Verdict.Rejected,
                OperationCanceledException => Verdict.Canceled,
                _ => Verdict.Crash
            };
            string detail = exception is YamlException yaml ? yaml.Code + (exception is YamlLimitExceededException limit ? " " + limit.LimitName : string.Empty) : exception.GetType().FullName + ": " + FirstLine(exception.Message) + " at " + TopFrame(exception);
            return Bounded(new StageResult(verdict, detail, watch.Elapsed, allocated, null), inputLength);
        }

        if (task.IsCanceled) return new StageResult(Verdict.Canceled, "canceled", watch.Elapsed, allocated, null);
        return Bounded(new StageResult(Verdict.Accepted, task.Result, watch.Elapsed, allocated, task.Result), inputLength);
    }

    private static StageResult Bounded(StageResult result, int inputLength) {
        long budget = Math.Max(AllocationFloor, (long)inputLength * AllocationFactor);
        return result.Allocated > budget && !result.IsFailure
            ? new StageResult(Verdict.Memory, "allocated " + YamlLargeCorpus.Format(result.Allocated) + " bytes for " + YamlLargeCorpus.Format(inputLength) + " input bytes (" + result.Verdict + ")", result.Elapsed, result.Allocated, result.RoundTripDifference)
            : result;
    }

    internal static byte[] Mutate(byte[] input, MutationKind kind, Random random, out string description) {
        switch (kind) {
            case MutationKind.Truncate: {
                int length = input.Length == 0 ? 0 : random.Next(input.Length);
                description = "truncated to " + length + " bytes";
                return input.AsSpan(0, length).ToArray();
            }
            case MutationKind.ByteFlip: {
                byte[] output = (byte[])input.Clone();
                int flips = random.Next(1, 9);
                var positions = new List<string>();
                for (int i = 0; i < flips && output.Length > 0; i++) {
                    int position = random.Next(output.Length);
                    output[position] = (byte)random.Next(256);
                    positions.Add(position.ToString(CultureInfo.InvariantCulture));
                }

                description = "random bytes at " + string.Join(",", positions);
                return output;
            }
            case MutationKind.IndicatorInsert: {
                const string indicators = "-?:,[]{}#&*!|>'\"%@`\t \n\r";
                var inserts = new List<(int, byte[])>();
                int count = random.Next(1, 6);
                for (int i = 0; i < count; i++) inserts.Add((random.Next(input.Length + 1), new[] { (byte)indicators[random.Next(indicators.Length)] }));
                description = "inserted " + string.Join(",", inserts.Select(x => x.Item1 + ":" + Show(x.Item2)));
                return Insert(input, inserts);
            }
            case MutationKind.LineEdit: {
                List<byte[]> lines = SplitLines(input);
                int operations = random.Next(1, 4);
                var applied = new List<string>();
                for (int i = 0; i < operations && lines.Count > 0; i++) {
                    int a = random.Next(lines.Count);
                    int b = random.Next(lines.Count);
                    switch (random.Next(4)) {
                        case 0: lines.Insert(a, lines[a]); applied.Add("dup " + a); break;
                        case 1: lines.RemoveAt(a); applied.Add("del " + a); break;
                        case 2: (lines[a], lines[b]) = (lines[b], lines[a]); applied.Add("swap " + a + "," + b); break;
                        default: byte[] moved = lines[a]; lines.RemoveAt(a); lines.Insert(Math.Min(b, lines.Count), moved); applied.Add("move " + a + "->" + b); break;
                    }
                }

                description = "lines " + string.Join("; ", applied);
                return lines.SelectMany(l => l).ToArray();
            }
            case MutationKind.IndentShift: {
                List<byte[]> lines = SplitLines(input);
                int edits = random.Next(1, 4);
                var applied = new List<string>();
                for (int i = 0; i < edits && lines.Count > 0; i++) {
                    int index = random.Next(lines.Count);
                    byte[] line = lines[index];
                    int indent = 0;
                    while (indent < line.Length && line[indent] == (byte)' ') indent++;
                    int change = random.Next(3);
                    if (change == 0) lines[index] = Enumerable.Repeat((byte)' ', random.Next(1, 5)).Concat(line).ToArray();
                    else if (change == 1) lines[index] = line.Skip(Math.Min(indent, random.Next(1, 5))).ToArray();
                    else lines[index] = new[] { (byte)'\t' }.Concat(line.Skip(indent)).ToArray();
                    applied.Add(index + (change == 0 ? "+" : change == 1 ? "-" : "tab"));
                }

                description = "indentation " + string.Join(",", applied);
                return lines.SelectMany(l => l).ToArray();
            }
            case MutationKind.DeepFlowNesting: {
                int depth = new[] { 300, 5_000, 200_000 }[random.Next(3)];
                bool balanced = random.Next(2) == 0;
                bool mapping = random.Next(2) == 0;
                var text = new StringBuilder("deep: ");
                for (int i = 0; i < depth; i++) text.Append(mapping ? "{a: " : "[");
                text.Append('x');
                if (balanced) for (int i = 0; i < depth; i++) text.Append(mapping ? '}' : ']');
                text.Append('\n');
                description = (balanced ? "balanced " : "unclosed ") + (mapping ? "flow mappings" : "flow sequences") + " nested " + depth + " deep";
                return InsertAtLineStart(input, random, Encoding.UTF8.GetBytes(text.ToString()));
            }
            case MutationKind.DeepBlockNesting: {
                int depth = new[] { 300, 3_000, 50_000 }[random.Next(3)];
                var text = new StringBuilder();
                if (random.Next(2) == 0) {
                    text.Append("\n---\n");
                    for (int i = 0; i < depth; i++) text.Append("- ");
                    text.Append("x\n");
                    description = "compact block sequences nested " + depth + " deep";
                } else {
                    depth = Math.Min(depth, 3_000);
                    text.Append("\n---\n");
                    for (int i = 0; i < depth; i++) text.Append(' ', i).Append("k:\n");
                    text.Append(' ', depth).Append("x\n");
                    description = "indented block mappings nested " + depth + " deep";
                }

                return Concat(input, Encoding.UTF8.GetBytes(text.ToString()));
            }
            case MutationKind.AliasBomb: {
                int levels = random.Next(5, 13);
                int fanOut = random.Next(5, 13);
                var text = new StringBuilder("\n---\nl0: &l0 [");
                text.Append(string.Join(",", Enumerable.Repeat("lol", fanOut))).Append("]\n");
                for (int level = 1; level < levels; level++) {
                    text.Append('l').Append(level).Append(": &l").Append(level).Append(" [");
                    text.Append(string.Join(",", Enumerable.Repeat("*l" + (level - 1), fanOut))).Append("]\n");
                }

                if (random.Next(2) == 0) text.Append("merged:\n  <<: [*l" + (levels - 1) + "]\n");
                description = "alias bomb with " + levels + " levels of " + fanOut + " aliases";
                return Concat(input, Encoding.UTF8.GetBytes(text.ToString()));
            }
            case MutationKind.AnchorChaos: {
                string[] fragments = { "&a ", "*a ", "*undefined", "&a [*a]", "&x &y ", "*", "&", "!!binary ", "!<tag:x> ", "!e!x ", "<<: *a\n", "? *a\n: *a\n" };
                var inserts = new List<(int, byte[])>();
                int count = random.Next(1, 5);
                for (int i = 0; i < count; i++) inserts.Add((random.Next(input.Length + 1), Encoding.UTF8.GetBytes(fragments[random.Next(fragments.Length)])));
                description = "anchor/alias/tag fragments " + string.Join(",", inserts.Select(x => x.Item1 + ":" + Show(x.Item2)));
                return Insert(input, inserts);
            }
            case MutationKind.HugeScalar: {
                int size = new[] { 1 << 20, 8 << 20, 17 << 20 }[random.Next(3)];
                int style = random.Next(4);
                var text = new StringBuilder(size + 64).Append("\n---\nhuge: ");
                switch (style) {
                    case 0: while (text.Length < size) text.Append("lorem ipsum "); break;
                    case 1: text.Append('"'); while (text.Length < size) text.Append("esc\\n\\t\\u0041\\x42 "); text.Append('"'); break;
                    case 2: text.Append("|\n"); while (text.Length < size) text.Append("  line of literal text\n"); break;
                    default: text.Append('\''); while (text.Length < size) text.Append("it''s quoted "); text.Append('\''); break;
                }

                text.Append('\n');
                description = new[] { "plain", "double-quoted", "literal", "single-quoted" }[style] + " scalar of " + size + " characters";
                return Concat(input, Encoding.UTF8.GetBytes(text.ToString()));
            }
            case MutationKind.NodeFlood: {
                int nodes = new[] { 200_000, 1_100_000 }[random.Next(2)];
                var text = new StringBuilder(nodes * 2 + 16).Append("\n---\n[");
                for (int i = 0; i < nodes; i++) text.Append(i == 0 ? "1" : ",1");
                text.Append("]\n");
                description = "flow sequence of " + nodes + " scalars";
                return Concat(input, Encoding.UTF8.GetBytes(text.ToString()));
            }
            case MutationKind.InvalidUtf8: {
                byte[][] sequences = {
                    new byte[] { 0x80 }, new byte[] { 0xC0, 0xAF }, new byte[] { 0xED, 0xA0, 0x80 }, new byte[] { 0xF8, 0x88, 0x80, 0x80, 0x80 },
                    new byte[] { 0xE2, 0x82 }, new byte[] { 0xFF }, new byte[] { 0xF4, 0x90, 0x80, 0x80 }, new byte[] { 0xC2 }
                };
                var inserts = new List<(int, byte[])>();
                int count = random.Next(1, 5);
                for (int i = 0; i < count; i++) inserts.Add((random.Next(input.Length + 1), sequences[random.Next(sequences.Length)]));
                description = "invalid UTF-8 " + string.Join(",", inserts.Select(x => x.Item1 + ":" + Show(x.Item2)));
                return Insert(input, inserts);
            }
            case MutationKind.ControlCharacters: {
                byte[][] sequences = {
                    new byte[] { 0x00 }, new byte[] { 0x01 }, new byte[] { 0x7F }, new byte[] { 0x1B }, new byte[] { 0xEF, 0xBB, 0xBF },
                    new byte[] { (byte)'\r' }, new byte[] { 0xC2, 0x85 }, new byte[] { 0xE2, 0x80, 0xA8 }, new byte[] { 0xEF, 0xBF, 0xBE }
                };
                var inserts = new List<(int, byte[])>();
                int count = random.Next(1, 5);
                for (int i = 0; i < count; i++) inserts.Add((random.Next(input.Length + 1), sequences[random.Next(sequences.Length)]));
                description = "control characters " + string.Join(",", inserts.Select(x => x.Item1 + ":" + Show(x.Item2)));
                return Insert(input, inserts);
            }
            default: {
                string text = Encoding.UTF8.GetString(input);
                int choice = random.Next(5);
                description = new[] { "UTF-16LE without BOM", "UTF-16BE with BOM", "UTF-32LE with BOM", "UTF-8 bytes behind a UTF-16LE BOM", "UTF-16LE with an odd byte cut off" }[choice];
                return choice switch {
                    0 => new UnicodeEncoding(false, false).GetBytes(text),
                    1 => new UnicodeEncoding(true, true).GetPreamble().Concat(new UnicodeEncoding(true, false).GetBytes(text)).ToArray(),
                    2 => new UTF32Encoding(false, true).GetPreamble().Concat(new UTF32Encoding(false, false).GetBytes(text)).ToArray(),
                    3 => new byte[] { 0xFF, 0xFE }.Concat(input).ToArray(),
                    _ => new UnicodeEncoding(false, true).GetPreamble().Concat(new UnicodeEncoding(false, false).GetBytes(text)).SkipLast(1).ToArray()
                };
            }
        }
    }

    private static string WriteReport(IReadOnlyList<MutantResult> results, int seed, int files, int perKind, TimeSpan timeout, TimeSpan elapsed) {
        var text = new StringBuilder();
        text.AppendLine("# YAML mutation run");
        text.AppendLine();
        text.AppendLine("- Seed " + seed + ", " + files + " corpus files, " + perKind + " variants per kind, " + results.Count + " mutants, " + YamlLargeCorpus.Format(results.Sum(r => (long)r.Length)) + " bytes in total");
        text.AppendLine("- Per-stage timeout " + timeout.TotalSeconds.ToString(CultureInfo.InvariantCulture) + " s; allocation budget per stage max(" + YamlLargeCorpus.Format(AllocationFloor) + " bytes, " + AllocationFactor + " x input); " + RuntimeInformation.FrameworkDescription + ", " + RuntimeInformation.OSDescription);
        text.AppendLine("- Finished in " + elapsed.TotalSeconds.ToString("F0", CultureInfo.InvariantCulture) + " s; largest working set seen after a stage " + YamlLargeCorpus.Format(PeakWorkingSet) + " bytes");
        text.AppendLine("- Pass: only YamlException/YamlLimitExceededException or OperationCanceledException escape, no hang, allocation within budget. Crash, Hang and Memory are failures.");
        text.AppendLine();
        foreach (bool reader in new[] { true, false }) {
            text.AppendLine(reader ? "## Reader (YamlReader over a stream, every value read)" : "## Model (LoadAsync, unchanged Save, resolved view)");
            text.AppendLine();
            text.Append("| Kind |");
            foreach (Verdict verdict in Enum.GetValues<Verdict>()) text.Append(' ').Append(verdict).Append(" |");
            text.AppendLine(" Slowest ms | Max allocated MB |");
            text.Append("|---|");
            foreach (Verdict _ in Enum.GetValues<Verdict>()) text.Append("---|");
            text.AppendLine("---|---|");
            foreach (var group in results.GroupBy(r => r.Kind).OrderBy(g => g.Key)) {
                text.Append("| ").Append(group.Key).Append(" |");
                foreach (Verdict verdict in Enum.GetValues<Verdict>()) text.Append(' ').Append(group.Count(r => (reader ? r.Reader : r.Model).Verdict == verdict)).Append(" |");
                text.Append(' ').Append(group.Max(r => (reader ? r.Reader : r.Model).Elapsed.TotalMilliseconds).ToString("F0", CultureInfo.InvariantCulture)).Append(" |");
                text.Append(' ').Append((group.Max(r => (reader ? r.Reader : r.Model).Allocated) / 1048576.0).ToString("F0", CultureInfo.InvariantCulture)).AppendLine(" |");
            }

            text.AppendLine();
        }

        int mismatches = results.Count(r => r.RoundTripMismatch);
        int loaded = results.Count(r => r.Model.Verdict == Verdict.Accepted);
        text.AppendLine("## Unchanged save of loaded mutants");
        text.AppendLine();
        text.AppendLine((loaded - mismatches) + " of " + loaded + " mutants that loaded saved back byte for byte.");
        text.AppendLine();

        var failures = results.SelectMany(r => new[] { ("reader", r, r.Reader), ("model", r, r.Model) }).Where(x => x.Item3.IsFailure).ToList();
        text.AppendLine("## Failures");
        text.AppendLine();
        if (failures.Count == 0 && mismatches == 0) text.AppendLine("None.");
        foreach (var group in failures.GroupBy(f => f.Item1 + " " + f.Item3.Verdict + ": " + f.Item3.Detail).OrderByDescending(g => g.Count()).Take(40)) {
            text.AppendLine("- " + group.Key + " (" + group.Count() + " mutants, e.g. " + group.First().Item2.Id + ", " + group.First().Item2.Description + ")");
        }

        foreach (MutantResult mismatch in results.Where(r => r.RoundTripMismatch).Take(20)) text.AppendLine("- round-trip mismatch: " + mismatch.Id + ", " + mismatch.Description + ": " + mismatch.Model.Detail);
        text.AppendLine();
        return text.ToString();
    }

    private static int StableSeed(int seed, string name, MutationKind kind, int variant) {
        unchecked {
            uint hash = 2166136261;
            foreach (char c in name + "|" + kind + "|" + variant.ToString(CultureInfo.InvariantCulture)) hash = (hash ^ c) * 16777619;
            return (int)(hash ^ (uint)seed);
        }
    }

    private static byte[] Insert(byte[] input, List<(int Position, byte[] Bytes)> inserts) {
        var output = new List<byte>(input.Length + inserts.Sum(i => i.Bytes.Length));
        var ordered = inserts.OrderBy(i => i.Position).ToList();
        int next = 0;
        for (int position = 0; position <= input.Length; position++) {
            while (next < ordered.Count && ordered[next].Position == position) output.AddRange(ordered[next++].Bytes);
            if (position < input.Length) output.Add(input[position]);
        }

        return output.ToArray();
    }

    private static byte[] InsertAtLineStart(byte[] input, Random random, byte[] text) {
        var starts = new List<int> { 0 };
        for (int i = 0; i < input.Length; i++) {
            if (input[i] == (byte)'\n') starts.Add(i + 1);
        }

        int position = starts[random.Next(starts.Count)];
        return Insert(input, new List<(int, byte[])> { (position, text) });
    }

    private static byte[] Concat(byte[] input, byte[] suffix) {
        var output = new byte[input.Length + suffix.Length];
        Buffer.BlockCopy(input, 0, output, 0, input.Length);
        Buffer.BlockCopy(suffix, 0, output, input.Length, suffix.Length);
        return output;
    }

    private static List<byte[]> SplitLines(byte[] input) {
        var lines = new List<byte[]>();
        int start = 0;
        for (int i = 0; i < input.Length; i++) {
            if (input[i] != (byte)'\n') continue;
            lines.Add(input.AsSpan(start, i + 1 - start).ToArray());
            start = i + 1;
        }

        if (start < input.Length) lines.Add(input.AsSpan(start).ToArray());
        return lines;
    }

    private static string Show(byte[] bytes) {
        var text = new StringBuilder();
        foreach (byte b in bytes) text.Append(b >= 0x21 && b < 0x7F ? ((char)b).ToString() : "\\x" + b.ToString("X2", CultureInfo.InvariantCulture));
        return text.Length > 24 ? text.ToString(0, 21) + "..." : text.ToString();
    }

    private static string SafeName(string id) {
        var text = new StringBuilder(id.Length);
        foreach (char c in id) text.Append(char.IsLetterOrDigit(c) || c is '-' or '.' ? c : '_');
        return text.Length > 180 ? text.ToString(0, 180) : text.ToString();
    }

    private static string FirstLine(string message) {
        int newline = message.IndexOf('\n');
        return newline < 0 ? message : message.Substring(0, newline).TrimEnd('\r');
    }

    private static string TopFrame(Exception exception) {
        string? trace = exception.StackTrace;
        if (trace == null) return "?";
        string first = FirstLine(trace).Trim();
        return first.StartsWith("at ", StringComparison.Ordinal) ? first.Substring(3) : first;
    }

    private static string Tsv(string? text) => (text ?? string.Empty).Replace("\t", "\\t").Replace("\r", "\\r").Replace("\n", "\\n");

    private static string? Option(string[] args, string name) {
        for (int i = 0; i < args.Length - 1; i++) {
            if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase)) return args[i + 1];
        }

        return null;
    }

    private sealed record StageResult(Verdict Verdict, string? Detail, TimeSpan Elapsed, long Allocated, string? RoundTripDifference) {
        public bool IsFailure => Verdict is Verdict.Crash or Verdict.Hang or Verdict.Memory;
    }

    private sealed record MutantResult(string Id, YamlLargeCorpusFile File, MutationKind Kind, string Description, int Length, StageResult Reader, StageResult Model) {
        public bool RoundTripMismatch => Model.Verdict == Verdict.Accepted && Model.RoundTripDifference != null;
    }
}
