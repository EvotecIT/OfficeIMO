using System.Diagnostics;
using System.Globalization;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelDataTableExecutionPairedRunner {
    private const int WarmupIterations = 12;

    internal static void Run(string[] arguments) {
        int iterations = arguments.Length > 1 && int.TryParse(arguments[1], out int parsedIterations)
            ? parsedIterations
            : 40;
        if (iterations <= 0) {
            throw new ArgumentOutOfRangeException(nameof(iterations));
        }

        string affinity = arguments.Length > 2
            ? BenchmarkProcessorAffinity.Apply(arguments[2])
            : "unchanged";
        string priority = arguments.Length > 3
            && !string.Equals(arguments[3], "unchanged", StringComparison.OrdinalIgnoreCase)
                ? BenchmarkProcessorAffinity.ApplyPriority(arguments[3])
                : Process.GetCurrentProcess().PriorityClass.ToString();

        var benchmark = new ExcelDataTableExecutionBenchmarks { RowCount = 25_000 };
        benchmark.Setup();
        for (int index = 0; index < WarmupIterations; index++) {
            benchmark.Automatic();
            benchmark.Parallel();
        }

        var automaticSamples = new double[iterations];
        var parallelSamples = new double[iterations];
        var pairedRatios = new double[iterations];
        for (int index = 0; index < iterations; index++) {
            double automaticFirst;
            double automaticSecond;
            double parallelFirst;
            double parallelSecond;
            if ((index & 1) == 0) {
                automaticFirst = MeasureMilliseconds(benchmark.Automatic);
                parallelFirst = MeasureMilliseconds(benchmark.Parallel);
                parallelSecond = MeasureMilliseconds(benchmark.Parallel);
                automaticSecond = MeasureMilliseconds(benchmark.Automatic);
            } else {
                parallelFirst = MeasureMilliseconds(benchmark.Parallel);
                automaticFirst = MeasureMilliseconds(benchmark.Automatic);
                automaticSecond = MeasureMilliseconds(benchmark.Automatic);
                parallelSecond = MeasureMilliseconds(benchmark.Parallel);
            }

            automaticSamples[index] = (automaticFirst + automaticSecond) / 2d;
            parallelSamples[index] = (parallelFirst + parallelSecond) / 2d;
            pairedRatios[index] = parallelSamples[index] / automaticSamples[index];
        }

        double automaticMedian = Median(automaticSamples);
        double parallelMedian = Median(parallelSamples);
        double ratioMedian = Median(pairedRatios);
        Console.WriteLine(string.Format(
            CultureInfo.InvariantCulture,
            "Paired DataTable execution ({0} warmups, {1} ABBA samples, affinity {2}, priority {3}): " +
            "Automatic median {4:F3} ms, Parallel median {5:F3} ms, ratio of medians {6:F4}, " +
            "paired ratio median {7:F4} (P25 {8:F4}, P75 {9:F4}).",
            WarmupIterations,
            iterations,
            affinity,
            priority,
            automaticMedian,
            parallelMedian,
            parallelMedian / automaticMedian,
            ratioMedian,
            Percentile(pairedRatios, 0.25d),
            Percentile(pairedRatios, 0.75d)));
    }

    private static double MeasureMilliseconds(Func<int> operation) {
        long started = Stopwatch.GetTimestamp();
        _ = operation();
        return Stopwatch.GetElapsedTime(started).TotalMilliseconds;
    }

    private static double Median(double[] samples) {
        Array.Sort(samples);
        int middle = samples.Length / 2;
        return (samples.Length & 1) == 0
            ? (samples[middle - 1] + samples[middle]) / 2d
            : samples[middle];
    }

    private static double Percentile(double[] samples, double percentile) {
        Array.Sort(samples);
        double position = (samples.Length - 1) * percentile;
        int lower = (int)position;
        int upper = Math.Min(lower + 1, samples.Length - 1);
        double fraction = position - lower;
        return samples[lower] + (samples[upper] - samples[lower]) * fraction;
    }
}
