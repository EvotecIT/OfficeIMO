using System.Diagnostics;
using System.Text;

namespace OfficeIMO.Toml.Benchmarks.LargeCorpus;

/// <summary>
/// <c>--mutations [--seed n] [--sample n] [--kind k] [--file name]</c>: hostile variants of a seeded sample of corpus
/// files go through <see cref="TomlDocument.Load(Stream, TomlLoadOptions?)"/> with
/// <see cref="TomlLoadOptions.UntrustedDefaults"/> and through the stream reader. A mutant passes when only
/// <see cref="TomlException"/> (limits included) escapes, it finishes within the time limit, it allocates within budget
/// and, when it loads, it saves back byte for byte. Mutants depend only on the seed, file, kind and variant.
/// </summary>
internal static class TomlMutationRun {
    private const int Variants = 3;
    private static readonly TimeSpan TimeLimit = TimeSpan.FromSeconds(20);

    private static readonly string[] Kinds = {
        "truncate", "flip-bytes", "delete-line", "duplicate-line", "swap-lines", "invalid-utf8", "control-chars",
        "deep-arrays", "deep-inline-tables", "deep-dotted-key", "deep-header", "huge-string", "huge-multiline-string",
        "key-flood", "table-flood", "unclosed-multiline", "quote-runs", "lone-carriage-return"
    };

    public static int Run(string[] args) {
        string root = TomlLargeCorpus.ResolveRoot(TomlLargeCorpus.Option(args, "--cache"));
        string output = TomlLargeCorpus.Option(args, "--output-dir") ?? Path.Combine(root, "reports");
        int seed = int.Parse(TomlLargeCorpus.Option(args, "--seed") ?? "20261113");
        int sampleSize = int.Parse(TomlLargeCorpus.Option(args, "--sample") ?? "400");
        string? onlyKind = TomlLargeCorpus.Option(args, "--kind");
        string? onlyFile = TomlLargeCorpus.Option(args, "--file");
        List<TomlCorpusFile> files = TomlLargeCorpus.Files(root);
        List<TomlCorpusFile> sample = onlyFile is not null
            ? files.Where(f => f.Name == onlyFile).ToList()
            : Enumerable.Range(0, Math.Min(sampleSize, files.Count)).Select(i => files[(int)((long)i * files.Count / Math.Min(sampleSize, files.Count))]).ToList();

        int mutants = 0, loaded = 0, rejected = 0, limited = 0;
        long maxAllocated = 0;
        double maxRatio = 0;
        TimeSpan slowest = TimeSpan.Zero;
        var failures = new List<string>();
        var failureDirectory = Path.Combine(output, "mutation-failures");
        foreach (TomlCorpusFile file in sample) {
            byte[] original = file.Read();
            foreach (string kind in Kinds) {
                if (onlyKind is not null && kind != onlyKind) continue;
                for (int variant = 0; variant < Variants; variant++) {
                    byte[] mutant = Mutate(original, kind, new Random(Hash(seed, file.Name, kind, variant)));
                    mutants++;
                    Outcome outcome = Exercise(mutant);
                    if (outcome.Elapsed > slowest) slowest = outcome.Elapsed;
                    maxAllocated = Math.Max(maxAllocated, outcome.Allocated);
                    maxRatio = Math.Max(maxRatio, (double)outcome.Allocated / Math.Max(1, mutant.Length));
                    if (outcome.Failure is null) {
                        if (outcome.Loaded) loaded++;
                        else if (outcome.Limited) limited++;
                        else rejected++;
                        continue;
                    }
                    string name = file.Source + "-" + Path.GetFileName(file.Path) + "-" + kind + "-" + variant + ".toml";
                    failures.Add(file.Name + "\t" + kind + "\t" + variant + "\t" + outcome.Failure);
                    Directory.CreateDirectory(failureDirectory);
                    File.WriteAllBytes(Path.Combine(failureDirectory, name), mutant);
                }
            }
        }
        var report = new StringBuilder();
        report.AppendLine("# TOML mutation run");
        report.AppendLine();
        report.AppendLine("Seed " + seed + ", " + sample.Count + " files, " + Kinds.Length + " kinds, " + Variants + " variants each: " + TomlLargeCorpus.Format(mutants) + " mutants.");
        report.AppendLine();
        report.AppendLine("| Outcome | Mutants |");
        report.AppendLine("| --- | ---: |");
        report.AppendLine("| Loaded and saved back byte for byte | " + TomlLargeCorpus.Format(loaded) + " |");
        report.AppendLine("| Rejected with TomlException | " + TomlLargeCorpus.Format(rejected) + " |");
        report.AppendLine("| Stopped by a limit | " + TomlLargeCorpus.Format(limited) + " |");
        report.AppendLine("| Failed (crash, hang, allocation or round-trip) | " + TomlLargeCorpus.Format(failures.Count) + " |");
        report.AppendLine();
        report.AppendLine("Slowest mutant " + slowest.TotalMilliseconds.ToString("F0") + " ms; most allocation " + TomlLargeCorpus.Format(maxAllocated) + " bytes; highest allocation per input byte " + maxRatio.ToString("F1") + ".");
        foreach (string failure in failures.Take(50)) report.AppendLine("- " + failure);
        Directory.CreateDirectory(output);
        File.WriteAllText(Path.Combine(output, "mutations.md"), report.ToString());
        Console.Write(report);
        return failures.Count == 0 ? 0 : 1;
    }

    private sealed class Outcome {
        public bool Loaded;
        public bool Limited;
        public string? Failure;
        public long Allocated;
        public TimeSpan Elapsed;
    }

    /// <summary>Loads and stream-reads one mutant on its own thread, with a time limit and an allocation budget.</summary>
    private static Outcome Exercise(byte[] mutant) {
        var outcome = new Outcome();
        // The model keeps the decoded text and one node per value; 48 bytes per input byte covers that with room.
        long budget = 48L * mutant.Length + 64L * 1024 * 1024;
        using var cancellation = new CancellationTokenSource(TimeLimit);
        var thread = new Thread(() => {
            long before = GC.GetAllocatedBytesForCurrentThread();
            var watch = Stopwatch.StartNew();
            try {
                try {
                    using TomlDocument document = TomlDocument.LoadAsync(new MemoryStream(mutant), TomlLoadOptions.UntrustedDefaults, cancellation.Token).GetAwaiter().GetResult();
                    outcome.Loaded = true;
                    var saved = new MemoryStream();
                    document.Save(saved);
                    if (!saved.ToArray().AsSpan().SequenceEqual(mutant)) outcome.Failure = "loaded but saved different bytes";
                } catch (TomlLimitExceededException) {
                    outcome.Limited = true;
                } catch (TomlException) {
                }
                var options = new TomlReaderOptions { IncludeComments = true, MaxDepth = 64, MaxStringLength = 1024 * 1024, MaxInputBytes = 16L * 1024 * 1024 };
                using var reader = new TomlReader(new MemoryStream(mutant), options, cancellation.Token);
                try {
                    while (reader.Read()) { }
                } catch (TomlException) {
                }
            } catch (OperationCanceledException) {
                outcome.Failure = "did not finish within " + TimeLimit.TotalSeconds + " s";
            } catch (Exception ex) {
                outcome.Failure = ex.GetType().Name + ": " + ex.Message;
            }
            outcome.Elapsed = watch.Elapsed;
            outcome.Allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        }, 16 * 1024 * 1024);
        thread.Start();
        if (!thread.Join(TimeLimit + TimeLimit)) {
            outcome.Failure = "hung past cancellation";
            return outcome;
        }
        if (outcome.Failure is null && outcome.Allocated > budget) outcome.Failure = "allocated " + outcome.Allocated + " bytes, over the budget of " + budget;
        return outcome;
    }

    private static byte[] Mutate(byte[] original, string kind, Random random) {
        var bytes = new List<byte>(original);
        int Position() => bytes.Count == 0 ? 0 : random.Next(bytes.Count + 1);
        int LineStart() {
            int p = Position();
            while (p > 0 && bytes[p - 1] != (byte)'\n') p--;
            return p;
        }
        void Insert(string text) => bytes.InsertRange(LineStart(), Encoding.UTF8.GetBytes(text));
        string Repeat(string s, int n) => new StringBuilder(s.Length * n).Insert(0, s, n).ToString();
        List<byte[]> Lines() {
            var lines = new List<byte[]>();
            int start = 0;
            for (int i = 0; i < original.Length; i++) {
                if (original[i] == (byte)'\n') {
                    lines.Add(original.AsSpan(start, i + 1 - start).ToArray());
                    start = i + 1;
                }
            }
            if (start < original.Length) lines.Add(original.AsSpan(start).ToArray());
            return lines;
        }
        switch (kind) {
            case "truncate":
                return original.AsSpan(0, random.Next(original.Length + 1)).ToArray();
            case "flip-bytes":
                for (int i = 0, n = 1 + random.Next(8); i < n && bytes.Count > 0; i++) {
                    int p = random.Next(bytes.Count);
                    bytes[p] = (byte)(bytes[p] ^ (1 << random.Next(8)));
                }
                break;
            case "delete-line":
            case "duplicate-line":
            case "swap-lines": {
                List<byte[]> lines = Lines();
                if (lines.Count == 0) break;
                int a = random.Next(lines.Count), b = random.Next(lines.Count);
                if (kind == "delete-line") lines.RemoveAt(a);
                else if (kind == "duplicate-line") lines.Insert(a, lines[a]);
                else (lines[a], lines[b]) = (lines[b], lines[a]);
                return lines.SelectMany(l => l).ToArray();
            }
            case "invalid-utf8":
                bytes.InsertRange(Position(), new[] { new byte[] { 0xC3, 0x28 }, new byte[] { 0xED, 0xA0, 0x80 }, new byte[] { 0xF8 }, new byte[] { 0xE2, 0x82 } }[random.Next(4)]);
                break;
            case "control-chars":
                bytes.Insert(Position(), new byte[] { 0x00, 0x7F, 0x1B, 0x08, 0x0C }[random.Next(5)]);
                break;
            case "deep-arrays": {
                int depth = new[] { 63, 65, 100_000 }[random.Next(3)];
                Insert("deep = " + Repeat("[", depth) + Repeat("]", depth) + "\n");
                break;
            }
            case "deep-inline-tables": {
                int depth = new[] { 63, 65, 50_000 }[random.Next(3)];
                Insert("deep = " + Repeat("{ a = ", depth) + "1" + Repeat(" }", depth) + "\n");
                break;
            }
            case "deep-dotted-key":
                Insert(Repeat("k.", new[] { 63, 65, 200_000 }[random.Next(3)]) + "k = 1\n");
                break;
            case "deep-header":
                Insert("[" + Repeat("h.", new[] { 63, 65, 200_000 }[random.Next(3)]) + "h]\n");
                break;
            case "huge-string":
                Insert("huge = \"" + new string('x', new[] { 1024 * 1024 - 2, 1024 * 1024 + 1, 20 * 1024 * 1024 }[random.Next(3)]) + "\"\n");
                break;
            case "huge-multiline-string":
                Insert("huge = '''\n" + Repeat("line of text\n", new[] { 1_000, 100_000, 2_000_000 }[random.Next(3)]) + "'''\n");
                break;
            case "key-flood": {
                var flood = new StringBuilder("[flood]\n");
                for (int i = 0, n = new[] { 1_000, 210_000, 400_000 }[random.Next(3)]; i < n; i++) flood.Append('k').Append(i).Append(" = ").Append(i).Append('\n');
                bytes.AddRange(Encoding.UTF8.GetBytes(flood.ToString()));
                break;
            }
            case "table-flood": {
                var flood = new StringBuilder();
                for (int i = 0, n = new[] { 1_000, 110_000, 300_000 }[random.Next(3)]; i < n; i++) flood.Append("[[flood]]\n");
                bytes.AddRange(Encoding.UTF8.GetBytes(flood.ToString()));
                break;
            }
            case "unclosed-multiline":
                Insert(new[] { "open = \"\"\"\n", "open = '''\n", "open = [\n", "open = {\n" }[random.Next(4)]);
                break;
            case "quote-runs":
                Insert(new[] { "q = \"\"\"\"\"\"\"\"\n", "q = ''''''''\n", "q = \"\"\"a\"\"\"\"\"\n", "q = '''''\n" }[random.Next(4)]);
                break;
            case "lone-carriage-return":
                bytes.Insert(Position(), (byte)'\r');
                break;
        }
        return bytes.ToArray();
    }

    private static int Hash(int seed, string file, string kind, int variant) {
        unchecked {
            int hash = seed;
            foreach (char ch in file) hash = hash * 31 + ch;
            foreach (char ch in kind) hash = hash * 31 + ch;
            return hash * 31 + variant;
        }
    }
}
