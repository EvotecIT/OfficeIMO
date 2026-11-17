using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf.Benchmarks.TextExtraction;

/// <summary>
/// Aggregates result CSVs into per-suite tables: mean, p50, p99 and max over successful files, pass rate, errors,
/// timeouts and crashes. Pass means the engine returned text without an exception, timeout or process crash.
/// </summary>
internal static class ResultSummary {
    internal static int Run(string[] args) {
        string output = Path.GetFullPath(Options.Get(args, "--output") ?? throw new ArgumentException("--output is required."));
        string pattern = Options.Get(args, "--pattern") ?? "results-*.csv";
        // Shards with different counts can cover the same file; each (mode, suite, file, engine) counts once and
        // the most recently written result wins.
        var latest = new Dictionary<(string Mode, string Suite, string File, string Engine), Row>();
        int rowCount = 0;
        int skipped = 0;
        foreach (string path in Directory.EnumerateFiles(output, pattern).OrderBy(static path => File.GetLastWriteTimeUtc(path))) {
            foreach (string[] fields in ReadCsv(path).Skip(1)) {
                // A cancelled or still-running shard can leave its last row cut short.
                if (fields.Length < 7
                    || !double.TryParse(fields[5], NumberStyles.Float, CultureInfo.InvariantCulture, out double ms)
                    || !double.TryParse(fields[6], NumberStyles.Float, CultureInfo.InvariantCulture, out double wallMs)) {
                    skipped++;
                    continue;
                }
                var row = new Row(fields[0], fields[1], fields[2], fields[3], fields[4], ms, wallMs);
                latest[(row.Mode, row.Suite, row.File, row.Engine)] = row;
                rowCount++;
            }
        }
        if (skipped > 0) {
            Console.Error.WriteLine($"{skipped} malformed or truncated rows were skipped.");
        }
        if (rowCount > latest.Count) {
            Console.Error.WriteLine($"{rowCount - latest.Count} duplicate rows from overlapping shards were counted once, using the newest result file.");
        }
        List<Row> rows = latest.Values.ToList();

        var report = new StringBuilder();
        var metrics = new StringBuilder("mode,suite,engine,files,pass,errors,timeouts,crashes,pass_rate,mean_ms,p50_ms,p99_ms,max_ms,total_s,mean_wall_ms\n");
        foreach (var modeGroup in rows.GroupBy(static row => row.Mode).OrderBy(static group => group.Key, StringComparer.Ordinal)) {
            foreach (var suiteGroup in modeGroup.GroupBy(static row => row.Suite).OrderBy(static group => group.Key, StringComparer.Ordinal)) {
                report.AppendLine($"### {suiteGroup.Key} ({modeGroup.Key})");
                report.AppendLine();
                report.AppendLine("| Engine | Files | Pass rate | Mean ms | p50 ms | p99 ms | Max ms | Errors | Timeouts | Crashes |");
                report.AppendLine("| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |");
                foreach (var engineGroup in suiteGroup.GroupBy(static row => row.Engine).OrderBy(static group => group.Key, StringComparer.Ordinal)) {
                    int files = engineGroup.Count();
                    double[] times = engineGroup.Where(static row => row.Status == "ok").Select(static row => row.Ms).OrderBy(static ms => ms).ToArray();
                    int errors = engineGroup.Count(static row => row.Status == "error");
                    int timeouts = engineGroup.Count(static row => row.Status == "timeout");
                    int crashes = engineGroup.Count(static row => row.Status == "crash");
                    double passRate = 100.0 * times.Length / files;
                    double mean = times.Length == 0 ? 0 : times.Average();
                    double p50 = Percentile(times, 0.50), p99 = Percentile(times, 0.99), max = times.Length == 0 ? 0 : times[^1];
                    double wall = engineGroup.Average(static row => row.WallMs);
                    report.AppendLine(string.Create(CultureInfo.InvariantCulture,
                        $"| {engineGroup.Key} | {files} | {passRate:0.0}% | {mean:0.00} | {p50:0.00} | {p99:0.0} | {max:0} | {errors} | {timeouts} | {crashes} |"));
                    metrics.AppendLine(string.Create(CultureInfo.InvariantCulture,
                        $"{CorpusRunner.Csv(modeGroup.Key)},{CorpusRunner.Csv(suiteGroup.Key)},{CorpusRunner.Csv(engineGroup.Key)},{files},{times.Length},{errors},{timeouts},{crashes},{passRate:0.00},{mean:0.000},{p50:0.000},{p99:0.000},{max:0.000},{times.Sum() / 1000:0.000},{wall:0.000}"));
                }
                report.AppendLine();
            }
        }
        File.WriteAllText(Path.Combine(output, "summary.md"), report.ToString());
        File.WriteAllText(Path.Combine(output, "metrics.csv"), metrics.ToString());
        Console.Write(report);
        return 0;
    }

    private sealed record Row(string Suite, string File, string Engine, string Mode, string Status, double Ms, double WallMs);

    /// <summary>Nearest-rank percentile.</summary>
    private static double Percentile(double[] sorted, double fraction) {
        if (sorted.Length == 0) {
            return 0;
        }
        int rank = (int)Math.Ceiling(fraction * sorted.Length);
        return sorted[Math.Clamp(rank - 1, 0, sorted.Length - 1)];
    }

    internal static IEnumerable<string[]> ReadCsv(string path) {
        foreach (string line in File.ReadLines(path)) {
            var fields = new List<string>();
            var field = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < line.Length; i++) {
                char c = line[i];
                if (quoted) {
                    if (c == '"' && i + 1 < line.Length && line[i + 1] == '"') {
                        field.Append('"');
                        i++;
                    } else if (c == '"') {
                        quoted = false;
                    } else {
                        field.Append(c);
                    }
                } else if (c == '"') {
                    quoted = true;
                } else if (c == ',') {
                    fields.Add(field.ToString());
                    field.Clear();
                } else {
                    field.Append(c);
                }
            }
            fields.Add(field.ToString());
            yield return fields.ToArray();
        }
    }
}
