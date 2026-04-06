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

    internal static void PrepareForMeasurement() {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
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
