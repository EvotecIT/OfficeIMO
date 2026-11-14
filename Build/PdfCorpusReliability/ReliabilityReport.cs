using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OfficeIMO.PdfCorpusReliability;

/// <summary>Aggregates streamed per-file results into results.json and summary.md.</summary>
internal static partial class ReliabilityReport {
    private const int MarkdownListLimit = 60;

    public static IEnumerable<FileResult> ReadResults(string path) {
        string[] lines = File.ReadAllLines(path);
        int last = Array.FindLastIndex(lines, static line => line.Length != 0);
        for (int index = 0; index <= last; index++) {
            if (lines[index].Length == 0) continue;
            FileResult? result;
            try {
                result = JsonSerializer.Deserialize<FileResult>(lines[index], ReliabilityJson.Line);
            } catch (JsonException) when (index == last) {
                // The run stopped while writing its last record; earlier records are complete.
                yield break;
            }
            if (result is not null) yield return Reclassify(result);
        }
    }

    /// <summary>Removes a final record that was cut off before its line ended.</summary>
    public static void TrimIncompleteTail(string path) {
        byte[] bytes = File.ReadAllBytes(path);
        int end = Array.LastIndexOf(bytes, (byte)'\n') + 1;
        if (end == bytes.Length) return;
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Write);
        stream.SetLength(end);
    }

    /// <summary>Applies the current classification to results written by an earlier runner build.</summary>
    private static FileResult Reclassify(FileResult result) {
        if (result.Worker is null || result.ProcessFailure is not null) return result;
        foreach (OperationResult op in result.Worker.Operations) {
            if (op.ExceptionType == "OfficeIMO.Pdf.PdfMutationBlockedException") op.Outcome = Outcome.Blocked;
        }
        result.Outcome = Outcome.Worst(result.Worker.Operations);
        return result;
    }

    public static void Write(RunOptions options, int missing = 0) {
        // The last line wins when a resumed run repeats a file.
        List<FileResult> results = ReadResults(options.ResultsLinesPath)
            .GroupBy(static result => result.Path, StringComparer.Ordinal)
            .Select(static group => group.Last())
            .OrderBy(static result => result.Path, StringComparer.Ordinal)
            .ToList();

        var operationTotals = new SortedDictionary<string, SortedDictionary<string, int>>(StringComparer.Ordinal);
        foreach (FileResult result in results) {
            foreach (string op in Operations.All) {
                string outcome = OperationOutcome(result, op);
                if (!operationTotals.TryGetValue(op, out SortedDictionary<string, int>? counts)) operationTotals[op] = counts = new(StringComparer.Ordinal);
                counts[outcome] = counts.GetValueOrDefault(outcome) + 1;
            }
        }

        var collectionTotals = results
            .GroupBy(static result => result.Collection, StringComparer.Ordinal)
            .OrderBy(static group => group.Key, StringComparer.Ordinal)
            .ToDictionary(static group => group.Key,
                static group => group.GroupBy(static result => result.Outcome).OrderBy(static g => g.Key, StringComparer.Ordinal).ToDictionary(static g => g.Key, static g => g.Count()));

        List<FailureClass> classes = BuildFailureClasses(results);
        List<FileResult> processFailures = results.Where(static result => result.ProcessFailure is not null)
            .OrderByDescending(static result => Outcome.Severity(result.Outcome)).ThenBy(static result => result.Path, StringComparer.Ordinal).ToList();
        List<FileResult> qpdfAccepted = results.Where(static result => result.QpdfExitCode is 0 or 3 && IsOurFailure(result)).ToList();
        List<FileResult> mismatches = results.Where(static result => result.Worker?.Operations.Any(IsRoundTripMismatch) == true).ToList();
        List<FileResult> slowest = results.OrderByDescending(static result => result.WallMilliseconds).Take(40).ToList();
        List<FileResult> heaviest = results.OrderByDescending(static result => result.PeakWorkingSetBytes).Take(25).ToList();
        long[] walls = results.Select(static result => result.WallMilliseconds).OrderBy(static value => value).ToArray();

        var summary = new {
            generatedUtc = DateTime.UtcNow,
            corpus = options.CorpusRoot,
            metadata = options.MetadataRoot,
            collections = options.Collections,
            timeoutSeconds = options.Timeout.TotalSeconds,
            memoryLimitMb = options.MemoryLimitBytes / (1024 * 1024),
            files = results.Count,
            missingOnDisk = missing,
            sha256Mismatches = results.Count(static result => !result.Sha256Verified && result.Worker is not null),
            wallMilliseconds = new { p50 = Percentile(walls, 0.50), p95 = Percentile(walls, 0.95), p99 = Percentile(walls, 0.99), max = walls.LastOrDefault() },
            outcomesByCollection = collectionTotals,
            operations = operationTotals,
            failureClasses = classes,
            processFailures = processFailures.Select(result => Brief(result, options)),
            qpdfAcceptedButFailed = qpdfAccepted.Select(result => Brief(result, options)),
            roundTripMismatches = mismatches.Select(result => Brief(result, options)),
            slowest = slowest.Select(result => Brief(result, options)),
            results
        };
        File.WriteAllText(options.ResultsJsonPath, JsonSerializer.Serialize(summary, ReliabilityJson.Indented));
        File.WriteAllText(options.SummaryPath, Markdown(options, results, missing, collectionTotals, operationTotals, classes, processFailures, qpdfAccepted, mismatches, slowest, heaviest, walls));
    }

    private static bool IsRoundTripMismatch(OperationResult op) =>
        op.Name is Operations.RoundTrip or Operations.Rewrite && op.Outcome == Outcome.Mismatch;

    internal static string OperationOutcome(FileResult result, string op) {
        OperationResult? operation = result.Worker?.Operations.FirstOrDefault(candidate => candidate.Name == op);
        if (operation is not null) return operation.Outcome;
        if (result.ProcessFailure is null) return Outcome.Skipped;
        if (result.InFlightOperation == op) return result.Outcome;
        // Operations before the one in flight ran, but the lost worker never reported their outcome.
        int inFlight = Array.IndexOf(Operations.All, result.InFlightOperation);
        return inFlight > Array.IndexOf(Operations.All, op) ? "unavailable" : "not reached";
    }

    /// <summary>Anything worse than a reasoned rejection, or any rejection other than a failed password check.</summary>
    private static bool IsOurFailure(FileResult result) {
        if (Outcome.Severity(result.Outcome) >= Outcome.Severity(Outcome.Mismatch)) return true;
        IReadOnlyList<OperationResult>? operations = result.Worker?.Operations;
        if (operations is null) return false;
        bool passwordRejected = operations.Any(static op => op.Name == Operations.Load && op.Outcome == Outcome.Rejected &&
            op.ExceptionType is "OfficeIMO.Pdf.PdfInvalidPasswordException" or "OfficeIMO.Pdf.PdfPasswordRequiredException");
        return !passwordRejected && operations.Any(static op => op.Outcome == Outcome.Rejected);
    }

    private static List<FailureClass> BuildFailureClasses(List<FileResult> results) {
        var classes = new Dictionary<string, FailureClass>(StringComparer.Ordinal);
        void Add(string op, string outcome, string? type, string? frame, string? message, FileResult result) {
            string normalized = NormalizeMessage(message);
            string key = string.Join("|", op, outcome, type, frame, normalized);
            if (!classes.TryGetValue(key, out FailureClass? failure)) {
                classes[key] = failure = new FailureClass { Operation = op, Outcome = outcome, ExceptionType = type, Frame = frame, Message = normalized };
            }
            failure.Count++;
            if (result.QpdfExitCode is 0 or 3) failure.QpdfAcceptedCount++;
            if (failure.Examples.Count < 3) failure.Examples.Add(result.Path);
        }
        foreach (FileResult result in results) {
            if (result.ProcessFailure is not null) {
                Add(result.InFlightOperation ?? "unknown", result.Outcome, null, null, ProcessFailureKind(result.ProcessFailure), result);
                continue;
            }
            // Later operations usually re-hit the same parse failure; count each distinct failure once per file,
            // under the first operation that reached it.
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (OperationResult op in result.Worker?.Operations ?? new List<OperationResult>()) {
                if (op.Outcome is Outcome.Ok or Outcome.Skipped) continue;
                string message = op.ExceptionType is null ? NormalizeMessage(op.Detail) : NormalizeMessage(op.Message);
                string signature = op.ExceptionType is null ? op.Name + "|" + op.Outcome : string.Join("|", op.ExceptionType, op.Frame, message);
                if (!seen.Add(signature)) continue;
                Add(op.Name, op.Outcome, op.ExceptionType, op.Frame, op.ExceptionType is null ? op.Detail : op.Message, result);
            }
        }
        return classes.Values
            .OrderByDescending(static failure => Outcome.Severity(failure.Outcome) >= 3 ? 1 : 0)
            .ThenByDescending(static failure => failure.Count)
            .ToList();
    }

    private static string ProcessFailureKind(string failure) {
        int frames = failure.IndexOf(" frames: ", StringComparison.Ordinal);
        if (frames > 0) {
            // Exit code plus the innermost OfficeIMO frame identifies a crash class.
            string code = failure[..failure.IndexOf(':')];
            string innermost = failure[(frames + 9)..].Split(" <- ")[0];
            return code + " at " + innermost;
        }
        int colon = failure.IndexOf(':');
        string head = colon > 0 ? failure[..colon] : failure;
        return NumberPattern().Replace(head, "N");
    }

    internal static string NormalizeMessage(string? message) {
        if (string.IsNullOrEmpty(message)) return "";
        string first = message.Split('\n')[0].Trim();
        first = QuotedPattern().Replace(first, "'…'");
        first = NumberPattern().Replace(first, "N");
        return first.Length <= 140 ? first : first[..140] + "…";
    }

    private static object Brief(FileResult result, RunOptions options) => new {
        result.Path,
        result.Sha256,
        result.Bytes,
        result.Collection,
        result.Source,
        result.QpdfExitCode,
        result.Outcome,
        result.ProcessFailure,
        result.InFlightOperation,
        result.WallMilliseconds,
        peakWorkingSetMb = result.PeakWorkingSetBytes / (1024 * 1024),
        failures = result.Worker?.Operations.Where(static op => op.Outcome is not (Outcome.Ok or Outcome.Skipped))
            .Select(static op => $"{op.Name}: {op.Outcome} {op.ExceptionType} {op.Frame} {op.Message ?? op.Detail}".Trim()),
        repro = Repro(result, options)
    };

    internal static string Repro(FileResult result, RunOptions options) {
        var builder = new StringBuilder("dotnet run -c Release --project Build/PdfCorpusReliability -- run");
        builder.Append(" --corpus \"").Append(options.CorpusRoot).Append('"');
        if (options.MetadataRoot is not null) builder.Append(" --metadata \"").Append(options.MetadataRoot).Append('"');
        builder.Append(" --collections ").Append(result.Collection);
        builder.Append(" --only \"").Append(result.Path).Append('"');
        builder.Append(" --corpus-version ").Append(options.CorpusVersion);
        builder.Append(" --timeout-seconds ").Append(options.Timeout.TotalSeconds.ToString("0", CultureInfo.InvariantCulture));
        builder.Append(" --memory-limit-mb ").Append((options.MemoryLimitBytes / (1024 * 1024)).ToString(CultureInfo.InvariantCulture));
        builder.Append(" --render-pages ").Append(options.RenderPages.ToString(CultureInfo.InvariantCulture));
        builder.Append(" --output \"").Append(Path.Combine(options.OutputDirectory, "repro")).Append('"');
        return builder.ToString();
    }

    private static long Percentile(long[] sorted, double percentile) =>
        sorted.Length == 0 ? 0 : sorted[Math.Min(sorted.Length - 1, (int)Math.Ceiling(percentile * sorted.Length) - 1)];

    private static string Markdown(RunOptions options, List<FileResult> results, int missing,
        Dictionary<string, Dictionary<string, int>> collectionTotals,
        SortedDictionary<string, SortedDictionary<string, int>> operationTotals,
        List<FailureClass> classes, List<FileResult> processFailures, List<FileResult> qpdfAccepted,
        List<FileResult> mismatches, List<FileResult> slowest, List<FileResult> heaviest, long[] walls) {
        var md = new StringBuilder();
        string[] outcomes = { Outcome.Ok, Outcome.Blocked, Outcome.Rejected, Outcome.Mismatch, Outcome.Unexpected, Outcome.Memory, Outcome.Timeout, Outcome.Crash, Outcome.Skipped, "not reached" };
        md.AppendLine("# PDF corpus reliability summary").AppendLine();
        md.AppendLine($"- Corpus: `{options.CorpusRoot}` (collections: {string.Join(", ", options.Collections)})");
        md.AppendLine($"- Files: {results.Count}; manifest entries missing on disk: {missing}; SHA-256 mismatches: {results.Count(static r => !r.Sha256Verified && r.Worker is not null)}");
        md.AppendLine($"- Limits: {options.Timeout.TotalSeconds:0} s per file, {options.MemoryLimitBytes / (1024 * 1024)} MB working set per worker, {options.Parallelism} workers");
        md.AppendLine($"- Wall time per file: p50 {Percentile(walls, .5)} ms, p95 {Percentile(walls, .95)} ms, p99 {Percentile(walls, .99)} ms, max {walls.LastOrDefault()} ms");
        md.AppendLine();

        md.AppendLine("## Worst outcome per file by collection").AppendLine();
        md.AppendLine("| Collection | " + string.Join(" | ", outcomes.Take(8)) + " |");
        md.AppendLine("|---|" + string.Concat(Enumerable.Repeat("---:|", 8)));
        foreach (var (collection, counts) in collectionTotals) {
            md.AppendLine($"| {collection} | " + string.Join(" | ", outcomes.Take(8).Select(o => counts.GetValueOrDefault(o).ToString(CultureInfo.InvariantCulture))) + " |");
        }
        md.AppendLine();

        md.AppendLine("## Outcomes per operation").AppendLine();
        md.AppendLine("| Operation | " + string.Join(" | ", outcomes) + " |");
        md.AppendLine("|---|" + string.Concat(Enumerable.Repeat("---:|", outcomes.Length)));
        foreach (string op in Operations.All) {
            SortedDictionary<string, int> counts = operationTotals.GetValueOrDefault(op) ?? new();
            md.AppendLine($"| {op} | " + string.Join(" | ", outcomes.Select(o => counts.GetValueOrDefault(o).ToString(CultureInfo.InvariantCulture))) + " |");
        }
        md.AppendLine();

        md.AppendLine("## Failure classes").AppendLine();
        md.AppendLine("Defect-like classes (unexpected, mismatch, memory, timeout, crash) first, then rejections. `qpdf ok` counts files qpdf checks as valid or warning-only.").AppendLine();
        md.AppendLine("| Count | qpdf ok | Operation | Outcome | Exception | First OfficeIMO frame | Message | Example |");
        md.AppendLine("|---:|---:|---|---|---|---|---|---|");
        foreach (FailureClass failure in classes.Take(MarkdownListLimit)) {
            md.AppendLine($"| {failure.Count} | {failure.QpdfAcceptedCount} | {failure.Operation} | {failure.Outcome} | {Short(failure.ExceptionType)} | {Cell(failure.Frame)} | {Cell(failure.Message)} | {Cell(failure.Examples.FirstOrDefault())} |");
        }
        md.AppendLine();

        md.AppendLine("## Crashes, timeouts and memory blow-ups").AppendLine();
        if (processFailures.Count == 0) md.AppendLine("None.");
        foreach (FileResult result in processFailures) {
            md.AppendLine($"- **{result.Outcome}** in `{result.InFlightOperation ?? "?"}`: `{result.Path}` ({result.Bytes:N0} bytes, qpdf {Qpdf(result.QpdfExitCode)}, peak {result.PeakWorkingSetBytes / (1024 * 1024)} MB, {result.WallMilliseconds} ms)");
            md.AppendLine($"  - SHA-256 `{result.Sha256}`; {Cell(result.ProcessFailure)}");
            md.AppendLine($"  - Repro: `{Repro(result, options)}`");
        }
        md.AppendLine();

        md.AppendLine($"## Files qpdf accepts that we reject or mishandle ({qpdfAccepted.Count})").AppendLine();
        md.AppendLine("| File | qpdf | Outcome | Failing operations |");
        md.AppendLine("|---|---:|---|---|");
        foreach (FileResult result in qpdfAccepted.OrderByDescending(static r => Outcome.Severity(r.Outcome)).Take(MarkdownListLimit * 2)) {
            md.AppendLine($"| {Cell(result.Path)} | {result.QpdfExitCode} | {result.Outcome} | {Cell(FailingOperations(result))} |");
        }
        if (qpdfAccepted.Count > MarkdownListLimit * 2) md.AppendLine().AppendLine($"{qpdfAccepted.Count - MarkdownListLimit * 2} more in results.json (`qpdfAcceptedButFailed`).");
        md.AppendLine();

        md.AppendLine($"## Round-trip mismatches ({mismatches.Count})").AppendLine();
        foreach (FileResult result in mismatches.Take(MarkdownListLimit)) {
            OperationResult op = result.Worker!.Operations.First(IsRoundTripMismatch);
            md.AppendLine($"- `{result.Path}` (qpdf {Qpdf(result.QpdfExitCode)}): {op.Detail}");
        }
        md.AppendLine();

        md.AppendLine("## Slowest files").AppendLine();
        md.AppendLine("| Wall ms | Peak MB | Bytes | Pages | Slowest operation | Outcome | File |");
        md.AppendLine("|---:|---:|---:|---:|---|---|---|");
        foreach (FileResult result in slowest) {
            OperationResult? op = result.Worker?.Operations.MaxBy(static op => op.Milliseconds);
            string slowOp = op is null ? result.InFlightOperation ?? "?" : $"{op.Name} {op.Milliseconds} ms";
            md.AppendLine($"| {result.WallMilliseconds} | {result.PeakWorkingSetBytes / (1024 * 1024)} | {result.Bytes:N0} | {result.Worker?.PageCount} | {slowOp} | {result.Outcome} | {Cell(result.Path)} |");
        }
        md.AppendLine();

        md.AppendLine("## Highest peak working set").AppendLine();
        md.AppendLine("| Peak MB | Bytes | Pages | Outcome | File |");
        md.AppendLine("|---:|---:|---:|---|---|");
        foreach (FileResult result in heaviest) {
            md.AppendLine($"| {result.PeakWorkingSetBytes / (1024 * 1024)} | {result.Bytes:N0} | {result.Worker?.PageCount} | {result.Outcome} | {Cell(result.Path)} |");
        }
        return md.ToString();
    }

    private static string FailingOperations(FileResult result) {
        if (result.ProcessFailure is not null) return $"{result.InFlightOperation}: {result.ProcessFailure}";
        return string.Join("; ", result.Worker?.Operations.Where(static op => op.Outcome is not (Outcome.Ok or Outcome.Skipped))
            .Select(static op => $"{op.Name} {op.Outcome} {Short(op.ExceptionType)} {NormalizeMessage(op.Message ?? op.Detail)}".Trim()) ?? Array.Empty<string>());
    }

    private static string Qpdf(int? code) => code switch { 0 => "ok", 3 => "warnings", 2 => "errors", null => "n/a", _ => code.Value.ToString(CultureInfo.InvariantCulture) };
    private static string Short(string? type) => type is null ? "" : type[(type.LastIndexOf('.') + 1)..];
    private static string Cell(string? value) => (value ?? "").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");

    [GeneratedRegex(@"\d+")]
    private static partial Regex NumberPattern();

    [GeneratedRegex(@"'[^']*'")]
    private static partial Regex QuotedPattern();

    private sealed class FailureClass {
        public string Operation { get; set; } = "";
        public string Outcome { get; set; } = "";
        public string? ExceptionType { get; set; }
        public string? Frame { get; set; }
        public string Message { get; set; } = "";
        public int Count { get; set; }
        public int QpdfAcceptedCount { get; set; }
        public List<string> Examples { get; } = new();
    }
}
