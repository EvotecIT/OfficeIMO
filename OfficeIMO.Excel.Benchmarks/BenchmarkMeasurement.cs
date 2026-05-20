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
        var allocated = new List<long>(measuredIterations);
        int lastMetric = 0;

        for (int i = 0; i < measuredIterations; i++) {
            PrepareForMeasurement();

            long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
            long startTimestamp = Stopwatch.GetTimestamp();
            lastMetric = action();
            long endTimestamp = Stopwatch.GetTimestamp();
            long allocatedAfter = GC.GetAllocatedBytesForCurrentThread();
            elapsed.Add(ElapsedMilliseconds(startTimestamp, endTimestamp));
            allocated.Add(Math.Max(0, allocatedAfter - allocatedBefore));
        }

        return new BenchmarkMeasurementResult(lastMetric, elapsed, allocated);
    }

    internal static IReadOnlyList<BenchmarkMeasurementResult> MeasureGroup(int warmupIterations, int measuredIterations, IReadOnlyList<Func<int>> actions) {
        if (actions == null) {
            throw new ArgumentNullException(nameof(actions));
        }

        if (actions.Count == 0) {
            return [];
        }

        var elapsed = new List<double>[actions.Count];
        var allocated = new List<long>[actions.Count];
        var outputMetrics = new int[actions.Count];
        for (int i = 0; i < actions.Count; i++) {
            if (actions[i] == null) {
                throw new ArgumentException("Benchmark actions cannot contain null entries.", nameof(actions));
            }

            elapsed[i] = new List<double>(measuredIterations);
            allocated[i] = new List<long>(measuredIterations);
        }

        for (int warmup = 0; warmup < warmupIterations; warmup++) {
            foreach (int index in GetRotatedOrder(actions.Count, warmup)) {
                actions[index]();
            }
        }

        for (int iteration = 0; iteration < measuredIterations; iteration++) {
            foreach (int index in GetRotatedOrder(actions.Count, iteration)) {
                PrepareForMeasurement();

                long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
                long startTimestamp = Stopwatch.GetTimestamp();
                outputMetrics[index] = actions[index]();
                long endTimestamp = Stopwatch.GetTimestamp();
                long allocatedAfter = GC.GetAllocatedBytesForCurrentThread();

                elapsed[index].Add(ElapsedMilliseconds(startTimestamp, endTimestamp));
                allocated[index].Add(Math.Max(0, allocatedAfter - allocatedBefore));
            }
        }

        var results = new BenchmarkMeasurementResult[actions.Count];
        for (int i = 0; i < actions.Count; i++) {
            results[i] = new BenchmarkMeasurementResult(outputMetrics[i], elapsed[i], allocated[i]);
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

    private static double ElapsedMilliseconds(long startTimestamp, long endTimestamp)
        => (endTimestamp - startTimestamp) * 1000.0 / Stopwatch.Frequency;

    internal sealed class BenchmarkMeasurementResult {
        internal BenchmarkMeasurementResult(
            int outputMetric,
            IReadOnlyList<double> samplesMilliseconds,
            IReadOnlyList<long> samplesAllocatedBytes) {
            OutputMetric = outputMetric;
            SamplesMilliseconds = samplesMilliseconds.ToArray();
            SamplesAllocatedBytes = samplesAllocatedBytes.ToArray();
        }

        internal int OutputMetric { get; }
        internal IReadOnlyList<double> SamplesMilliseconds { get; }
        internal IReadOnlyList<long> SamplesAllocatedBytes { get; }
        internal double AverageMilliseconds => SamplesMilliseconds.Count == 0 ? 0 : SamplesMilliseconds.Average();
        internal double StandardDeviationMilliseconds => CalculateStandardDeviation(SamplesMilliseconds);
        internal double StandardErrorMilliseconds => SamplesMilliseconds.Count == 0 ? 0 : StandardDeviationMilliseconds / Math.Sqrt(SamplesMilliseconds.Count);
        internal double AverageAllocatedBytes => SamplesAllocatedBytes.Count == 0 ? 0 : SamplesAllocatedBytes.Average();
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

        internal double MedianAllocatedBytes {
            get {
                if (SamplesAllocatedBytes.Count == 0) {
                    return 0;
                }

                var ordered = SamplesAllocatedBytes.OrderBy(v => v).ToArray();
                int middle = ordered.Length / 2;
                if ((ordered.Length & 1) == 1) {
                    return ordered[middle];
                }

                return (ordered[middle - 1] + ordered[middle]) / 2.0;
            }
        }

        private static double CalculateStandardDeviation(IReadOnlyList<double> values) {
            if (values.Count <= 1) {
                return 0;
            }

            double average = values.Average();
            double variance = values.Sum(value => Math.Pow(value - average, 2)) / (values.Count - 1);
            return Math.Sqrt(variance);
        }
    }
}
