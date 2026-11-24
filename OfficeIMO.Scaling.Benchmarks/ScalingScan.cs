using System.Diagnostics;
using System.Globalization;

namespace OfficeIMO.Scaling.Benchmarks;

/// <summary>
/// Quick stopwatch sweep for spotting super-linear lanes before a full
/// BenchmarkDotNet run. Prints time and allocation per unit at each size and the
/// growth of the per-unit cost from the smallest to the largest size; a linear
/// operation stays near 1x.
/// </summary>
internal static class ScalingScan {
    public static int Run(string[] args) {
        int[] sizes = [10, 100, 1000];
        string? filter = null;
        double budgetMs = 300;
        for (int i = 0; i < args.Length; i++) {
            switch (args[i]) {
                case "--sizes":
                    sizes = args[++i].Split(',').Select(value => int.Parse(value, CultureInfo.InvariantCulture)).ToArray();
                    break;
                case "--filter":
                    filter = args[++i];
                    break;
                case "--budget-ms":
                    budgetMs = double.Parse(args[++i], CultureInfo.InvariantCulture);
                    break;
            }
        }

        Console.WriteLine("| Lane | Unit | " + string.Join(" | ", sizes.Select(size => $"{size} µs/unit | {size} B/unit")) + " | time growth (last two sizes) | alloc growth |");
        Console.WriteLine("|---|---|" + string.Concat(Enumerable.Repeat("---:|---:|", sizes.Length)) + "---:|---:|");
        bool anyFailed = false;
        foreach (ScalingLane lane in ScalingLanes.Create()) {
            if (filter != null && lane.Name.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0) {
                continue;
            }

            var timePerUnit = new List<double>();
            var allocPerUnit = new List<double>();
            string? failure = null;
            foreach (int size in sizes) {
                try {
                    (double micros, double bytes) = Measure(lane, size, budgetMs);
                    timePerUnit.Add(micros / size);
                    allocPerUnit.Add(bytes / size);
                } catch (Exception exception) {
                    failure = $"{size}: {exception.GetType().Name}: {exception.Message}";
                    break;
                }
            }

            string cells = string.Join(" | ", timePerUnit.Select((time, index) => $"{time:0.0} | {allocPerUnit[index]:N0}"));
            if (failure != null) {
                anyFailed = true;
                Console.WriteLine($"| {lane.Name} | {lane.Unit} | {cells} | failed {failure} |");
                continue;
            }

            Console.WriteLine($"| {lane.Name} | {lane.Unit} | {cells} | {Growth(timePerUnit):0.0}x | {Growth(allocPerUnit):0.0}x |");
        }

        return anyFailed ? 1 : 0;
    }

    private static double Growth(List<double> perUnit) =>
        perUnit.Count < 2 ? 1 : perUnit[^1] / Math.Max(double.Epsilon, perUnit[^2]);

    private static (double Micros, double Bytes) Measure(ScalingLane lane, int size, double budgetMs) {
        lane.Setup(size);
        object? warm = lane.Run();
        lane.Validate(warm);
        GC.Collect();
        GC.WaitForPendingFinalizers();

        int iterations = 0;
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        object? last;
        do {
            last = lane.Run();
            iterations++;
        } while (stopwatch.Elapsed.TotalMilliseconds < budgetMs);

        stopwatch.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        lane.Validate(last);
        return (stopwatch.Elapsed.TotalMilliseconds * 1000 / iterations, (double)allocated / iterations);
    }
}
