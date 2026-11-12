using System.Diagnostics;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Yaml.Benchmarks.Kit;

/// <summary>
/// Runs the oneyaml-kit tasks on the kit's workloads:
/// <c>--kit-timings [--rounds n] [--task name] [--workload id,id] [--lane text] [--output file.md] [--validate-only] [--raise-limits]</c>.
/// Every lane is validated first. Then, per workload and task, the lanes run round-robin in rounds of about 50 ms
/// after a one-second warm-up each, so a change in machine load affects every lane alike; the table gives the median
/// time and the bytes allocated per operation. An edit consumes its document, so each edit run gets a document parsed
/// outside timing. A cross-check for <see cref="YamlKitBenchmarks"/> on a busy host.
/// </summary>
internal static class YamlKitTimings {
    public static int Run(string[] args) {
        int rounds = int.Parse(Option(args, "--rounds") ?? "11", CultureInfo.InvariantCulture);
        string? taskFilter = Option(args, "--task");
        string[]? workloadFilter = Option(args, "--workload")?.Split(',', StringSplitOptions.RemoveEmptyEntries);
        string? laneFilter = Option(args, "--lane");
        bool validateOnly = args.Contains("--validate-only");
        string? output = Option(args, "--output");
        if (args.Contains("--raise-limits")) YamlKitLanes.OfficeLoad = new YamlLoadOptions { MaxNodes = int.MaxValue };

        var text = new StringBuilder();
        text.AppendLine("| Task | Workload | Input | " + string.Join(" | ", YamlKitLanes.Libraries) + " |");
        text.AppendLine("|---|---|---:|" + string.Concat(YamlKitLanes.Libraries.Select(_ => "---:|")));
        Console.WriteLine(text.ToString().TrimEnd());
        var medians = new Dictionary<(YamlKitTask Task, string Library), List<(string Workload, double Micros)>>();
        int failures = 0;
        foreach (YamlKitTask task in Enum.GetValues<YamlKitTask>()) {
            if (taskFilter != null && !string.Equals(task.ToString(), taskFilter, StringComparison.OrdinalIgnoreCase)) continue;
            foreach (YamlKitWorkload workload in YamlKitWorkload.All) {
                if (!workload.Tasks.Contains(task) || (workloadFilter != null && !workloadFilter.Contains(workload.Id))) continue;
                var lanes = new List<YamlKitLane>();
                var cells = new Dictionary<string, string>();
                foreach (string library in YamlKitLanes.Libraries) {
                    if (laneFilter != null && !library.Contains(laneFilter, StringComparison.OrdinalIgnoreCase)) continue;
                    YamlKitLane? lane = YamlKitLanes.Find(library, task, workload);
                    if (lane == null) continue;
                    try {
                        lane.Validate(workload);
                        lanes.Add(lane);
                    } catch (Exception exception) {
                        failures++;
                        cells[library] = "fails";
                        Console.Error.WriteLine("FAIL " + task + " " + workload.Id + " " + library + ": " + Describe(exception));
                    }
                }

                if (!validateOnly && lanes.Count > 0) {
                    (double Median, long Allocated)[] results = Measure(lanes, workload, rounds);
                    for (int i = 0; i < lanes.Count; i++) {
                        cells[lanes[i].Library] = Format(results[i].Median) + " (" + Bytes(results[i].Allocated) + ")";
                        if (!medians.TryGetValue((task, lanes[i].Library), out var list)) medians[(task, lanes[i].Library)] = list = new();
                        list.Add((workload.Id, results[i].Median));
                    }
                } else {
                    foreach (YamlKitLane lane in lanes) cells[lane.Library] = "ok";
                }

                string row = "| " + Name(task) + " | " + workload.Id + " | " + Bytes(workload.Bytes.Length) + " | "
                    + string.Join(" | ", YamlKitLanes.Libraries.Select(l => cells.TryGetValue(l, out string? cell) ? cell : "-")) + " |";
                text.AppendLine(row);
                Console.WriteLine(row);
            }
        }

        if (!validateOnly) {
            text.AppendLine();
            text.AppendLine("Geometric mean of the medians over the workloads every listed lane completed, per task:");
            text.AppendLine();
            text.AppendLine("| Task | Workloads | " + string.Join(" | ", YamlKitLanes.Libraries) + " |");
            text.AppendLine("|---|---:|" + string.Concat(YamlKitLanes.Libraries.Select(_ => "---:|")));
            foreach (YamlKitTask task in Enum.GetValues<YamlKitTask>()) {
                var present = YamlKitLanes.Libraries.Where(l => medians.ContainsKey((task, l))).ToList();
                if (present.Count == 0) continue;
                // Compare on the workloads OfficeIMO and OneYaml both completed; other lanes show their own coverage.
                var common = medians.TryGetValue((task, "OfficeIMO"), out var ours) && medians.TryGetValue((task, "OneYaml"), out var theirs)
                    ? ours.Select(o => o.Workload).Intersect(theirs.Select(t => t.Workload)).ToHashSet()
                    : medians[(task, present[0])].Select(m => m.Workload).ToHashSet();
                text.Append("| ").Append(Name(task)).Append(" | ").Append(common.Count).Append(" | ");
                foreach (string library in YamlKitLanes.Libraries) {
                    if (!medians.TryGetValue((task, library), out var values)) {
                        text.Append("- | ");
                        continue;
                    }

                    var selected = values.Where(v => common.Contains(v.Workload)).ToList();
                    string cell = selected.Count == common.Count && selected.Count > 0
                        ? Format(Math.Exp(selected.Average(v => Math.Log(v.Micros))))
                        : selected.Count + "/" + common.Count + " workloads";
                    text.Append(cell).Append(" | ");
                }

                text.AppendLine();
            }

            Console.WriteLine();
            Console.WriteLine(text.ToString().Substring(text.ToString().IndexOf("Geometric mean", StringComparison.Ordinal)));
        }

        if (output != null) File.WriteAllText(output, text.ToString(), new UTF8Encoding(false));
        return failures == 0 ? 0 : 1;
    }

    /// <summary>The median time per operation in microseconds and the bytes allocated per operation, per lane.</summary>
    private static (double Median, long Allocated)[] Measure(IReadOnlyList<YamlKitLane> lanes, YamlKitWorkload workload, int rounds) {
        var samples = new double[lanes.Count][];
        var allocated = new long[lanes.Count];
        var iterations = new int[lanes.Count];
        var states = new object[lanes.Count];
        var watch = new Stopwatch();
        for (int lane = 0; lane < lanes.Count; lane++) {
            samples[lane] = new double[rounds];
            YamlKitLane current = lanes[lane];
            if (!current.ConsumesState) states[lane] = current.Prepare(workload);
            watch.Restart();
            int count = 0;
            while (watch.ElapsedMilliseconds < 1000) {
                RunOnce(current, workload, states[lane]);
                count++;
            }

            // About 50 ms per round, at least one run; edits keep a round's prepared documents within about 64 MB of input.
            iterations[lane] = Math.Max(1, (int)(count / 20.0));
            if (current.ConsumesState) iterations[lane] = Math.Max(1, Math.Min(iterations[lane], (int)(64L * 1024 * 1024 / Math.Max(1, workload.Bytes.Length))));
        }

        for (int round = 0; round < rounds; round++) {
            for (int lane = 0; lane < lanes.Count; lane++) {
                YamlKitLane current = lanes[lane];
                int count = iterations[lane];
                object[]? prepared = null;
                if (current.ConsumesState) {
                    prepared = new object[count];
                    for (int i = 0; i < count; i++) prepared[i] = current.Prepare(workload);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

                long before = GC.GetAllocatedBytesForCurrentThread();
                watch.Restart();
                if (prepared != null) {
                    for (int i = 0; i < count; i++) current.Run(prepared[i]);
                } else {
                    object state = states[lane];
                    for (int i = 0; i < count; i++) current.Run(state);
                }

                samples[lane][round] = watch.Elapsed.TotalMilliseconds * 1000 / count;
                allocated[lane] = (GC.GetAllocatedBytesForCurrentThread() - before) / count;
            }
        }

        var results = new (double, long)[lanes.Count];
        for (int lane = 0; lane < lanes.Count; lane++) {
            Array.Sort(samples[lane]);
            results[lane] = (samples[lane][rounds / 2], allocated[lane]);
        }

        return results;
    }

    private static void RunOnce(YamlKitLane lane, YamlKitWorkload workload, object? state) {
        lane.Run(lane.ConsumesState ? lane.Prepare(workload) : state!);
    }

    private static string Name(YamlKitTask task) => task switch {
        YamlKitTask.TypedLoad => "typed load",
        YamlKitTask.GenericLoad => "generic load",
        YamlKitTask.Dump => "dump",
        YamlKitTask.DocumentParse => "document parse",
        YamlKitTask.UnchangedEmit => "unchanged emit",
        _ => "edit script"
    };

    internal static string Format(double micros) => micros >= 1000
        ? (micros / 1000).ToString(micros >= 100_000 ? "N0" : "N2", CultureInfo.InvariantCulture) + " ms"
        : micros >= 1 ? micros.ToString(micros >= 100 ? "N0" : "N2", CultureInfo.InvariantCulture) + " µs"
        : (micros * 1000).ToString("N0", CultureInfo.InvariantCulture) + " ns";

    private static string Bytes(long bytes) => bytes >= 1024 * 1024
        ? (bytes / 1048576.0).ToString("N1", CultureInfo.InvariantCulture) + " MiB"
        : bytes >= 1024 ? (bytes / 1024.0).ToString("N1", CultureInfo.InvariantCulture) + " KiB" : bytes + " B";

    private static string Describe(Exception exception) {
        string message = exception.GetType().Name + ": " + exception.Message;
        return message.Length > 400 ? message.Substring(0, 400) + "..." : message;
    }

    private static string? Option(string[] args, string name) {
        for (int i = 0; i < args.Length - 1; i++) {
            if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase)) return args[i + 1];
        }

        return null;
    }
}
