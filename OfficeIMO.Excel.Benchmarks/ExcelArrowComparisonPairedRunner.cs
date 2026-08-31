using System.Diagnostics;
using System.Globalization;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelArrowComparisonPairedRunner {
    private const int WarmupIterations = 12;

    internal static void Run(string[] arguments) {
        string mode = arguments.Length > 1 ? arguments[1] : "explicit";
        int iterations = arguments.Length > 2 && int.TryParse(arguments[2], out int parsedIterations)
            ? parsedIterations
            : 30;
        if (iterations <= 0) {
            throw new ArgumentOutOfRangeException(nameof(iterations));
        }

        string affinity = arguments.Length > 3
            ? BenchmarkProcessorAffinity.Apply(arguments[3])
            : "unchanged";
        string priority = arguments.Length > 4
            && !string.Equals(arguments[4], "unchanged", StringComparison.OrdinalIgnoreCase)
                ? BenchmarkProcessorAffinity.ApplyPriority(arguments[4])
                : Process.GetCurrentProcess().PriorityClass.ToString();

        var benchmark = new ExcelArrowConversionBenchmarks();
        benchmark.Setup();
        Func<ArrowConversionObservation> officeIMO;
        Func<ArrowConversionObservation> excelReader;
        Action<ArrowConversionObservation> validate;
        if (string.Equals(mode, "explicit", StringComparison.OrdinalIgnoreCase)) {
            officeIMO = benchmark.OfficeIMO_ExplicitSchema;
            excelReader = benchmark.ExcelReaderNet_ExplicitSchema;
            validate = benchmark.ValidateExplicit;
        } else if (string.Equals(mode, "inferred", StringComparison.OrdinalIgnoreCase)) {
            officeIMO = benchmark.OfficeIMO_InferredSchema;
            excelReader = benchmark.ExcelReaderNet_InferredSchema;
            validate = benchmark.ValidateInferred;
            benchmark.EnsureInferredSchemaIsComparable();
        } else {
            throw new ArgumentException("Arrow comparison mode must be 'explicit' or 'inferred'.", nameof(arguments));
        }

        for (int index = 0; index < WarmupIterations; index++) {
            validate(officeIMO());
            validate(excelReader());
        }

        var officeSamples = new double[iterations];
        var excelReaderSamples = new double[iterations];
        var pairedRatios = new double[iterations];
        for (int index = 0; index < iterations; index++) {
            (double Milliseconds, ArrowConversionObservation Observation) officeFirst;
            (double Milliseconds, ArrowConversionObservation Observation) officeSecond;
            (double Milliseconds, ArrowConversionObservation Observation) excelReaderFirst;
            (double Milliseconds, ArrowConversionObservation Observation) excelReaderSecond;
            if ((index & 1) == 0) {
                officeFirst = Measure(officeIMO);
                excelReaderFirst = Measure(excelReader);
                excelReaderSecond = Measure(excelReader);
                officeSecond = Measure(officeIMO);
            } else {
                excelReaderFirst = Measure(excelReader);
                officeFirst = Measure(officeIMO);
                officeSecond = Measure(officeIMO);
                excelReaderSecond = Measure(excelReader);
            }

            validate(officeFirst.Observation);
            validate(officeSecond.Observation);
            validate(excelReaderFirst.Observation);
            validate(excelReaderSecond.Observation);
            officeSamples[index] = (officeFirst.Milliseconds + officeSecond.Milliseconds) / 2d;
            excelReaderSamples[index] = (excelReaderFirst.Milliseconds + excelReaderSecond.Milliseconds) / 2d;
            pairedRatios[index] = officeSamples[index] / excelReaderSamples[index];
        }

        double officeMedian = Median(officeSamples);
        double excelReaderMedian = Median(excelReaderSamples);
        Console.WriteLine(string.Format(
            CultureInfo.InvariantCulture,
            "Paired XLSX Arrow {0}-schema comparison ({1} warmups, {2} ABBA samples, affinity {3}, priority {4}): " +
            "OfficeIMO median {5:F3} ms, ExcelReader.NET median {6:F3} ms, ratio of medians {7:F4}, " +
            "paired ratio median {8:F4} (P25 {9:F4}, P75 {10:F4}).",
            mode.ToLowerInvariant(),
            WarmupIterations,
            iterations,
            affinity,
            priority,
            officeMedian,
            excelReaderMedian,
            officeMedian / excelReaderMedian,
            Median(pairedRatios),
            Percentile(pairedRatios, 0.25d),
            Percentile(pairedRatios, 0.75d)));
    }

    private static (double Milliseconds, ArrowConversionObservation Observation) Measure(
        Func<ArrowConversionObservation> operation) {
        long started = Stopwatch.GetTimestamp();
        ArrowConversionObservation observation = operation();
        return (Stopwatch.GetElapsedTime(started).TotalMilliseconds, observation);
    }

    private static double Median(double[] samples) {
        double[] ordered = [.. samples.OrderBy(static sample => sample)];
        int middle = ordered.Length / 2;
        return (ordered.Length & 1) == 0
            ? (ordered[middle - 1] + ordered[middle]) / 2d
            : ordered[middle];
    }

    private static double Percentile(double[] samples, double percentile) {
        double[] ordered = [.. samples.OrderBy(static sample => sample)];
        double position = (ordered.Length - 1) * percentile;
        int lower = (int)position;
        int upper = Math.Min(lower + 1, ordered.Length - 1);
        double fraction = position - lower;
        return ordered[lower] + (ordered[upper] - ordered[lower]) * fraction;
    }
}
