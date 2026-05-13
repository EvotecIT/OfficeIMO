using System.Diagnostics;

namespace OfficeIMO.Excel.Benchmarks;

internal static class BenchmarkMeasurement {
    internal static BenchmarkMeasurementResult Measure(int warmupIterations, int measuredIterations, Func<int> action) {
        if (action == null) {
            throw new ArgumentNullException(nameof(action));
        }

        for (int i = 0; i < warmupIterations; i++) {
            action();
        }

        var elapsed = new List<double>(measuredIterations);
        int lastMetric = 0;

        for (int i = 0; i < measuredIterations; i++) {
            PrepareForMeasurement();

            var stopwatch = Stopwatch.StartNew();
            lastMetric = action();
            stopwatch.Stop();
            elapsed.Add(stopwatch.Elapsed.TotalMilliseconds);
        }

        return new BenchmarkMeasurementResult(lastMetric, elapsed);
    }

    internal static IReadOnlyList<BenchmarkMeasurementResult> MeasureGroup(int warmupIterations, int measuredIterations, IReadOnlyList<Func<int>> actions) {
        if (actions == null) {
            throw new ArgumentNullException(nameof(actions));
        }

        if (actions.Count == 0) {
            return [];
        }

        var elapsed = new List<double>[actions.Count];
        var outputMetrics = new int[actions.Count];
        for (int i = 0; i < actions.Count; i++) {
            if (actions[i] == null) {
                throw new ArgumentException("Benchmark actions cannot contain null entries.", nameof(actions));
            }

            elapsed[i] = new List<double>(measuredIterations);
        }

        for (int warmup = 0; warmup < warmupIterations; warmup++) {
            foreach (int index in GetRotatedOrder(actions.Count, warmup)) {
                actions[index]();
            }
        }

        for (int iteration = 0; iteration < measuredIterations; iteration++) {
            foreach (int index in GetRotatedOrder(actions.Count, iteration)) {
                PrepareForMeasurement();

                var stopwatch = Stopwatch.StartNew();
                outputMetrics[index] = actions[index]();
                stopwatch.Stop();

                elapsed[index].Add(stopwatch.Elapsed.TotalMilliseconds);
            }
        }

        var results = new BenchmarkMeasurementResult[actions.Count];
        for (int i = 0; i < actions.Count; i++) {
            results[i] = new BenchmarkMeasurementResult(outputMetrics[i], elapsed[i]);
        }

        return results;
    }

    internal static void PrepareForMeasurement() {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    private static IEnumerable<int> GetRotatedOrder(int count, int offset) {
        for (int i = 0; i < count; i++) {
            yield return (i + offset) % count;
        }
    }

    internal sealed class BenchmarkMeasurementResult {
        internal BenchmarkMeasurementResult(int outputMetric, IReadOnlyList<double> samplesMilliseconds) {
            OutputMetric = outputMetric;
            SamplesMilliseconds = samplesMilliseconds.ToArray();
        }

        internal int OutputMetric { get; }
        internal IReadOnlyList<double> SamplesMilliseconds { get; }
        internal double AverageMilliseconds => SamplesMilliseconds.Count == 0 ? 0 : SamplesMilliseconds.Average();
        internal double MedianMilliseconds {
            get {
                if (SamplesMilliseconds.Count == 0) {
                    return 0;
                }

                var ordered = SamplesMilliseconds.OrderBy(v => v).ToArray();
                int middle = ordered.Length / 2;
                if ((ordered.Length & 1) == 1) {
                    return ordered[middle];
                }

                return (ordered[middle - 1] + ordered[middle]) / 2.0;
            }
        }
    }
}
