using System.Diagnostics;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.CSV.Benchmarks;

internal static class CsvExcelReaderComparisonPairedRunner {
    private const int WarmupIterations = 16;

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

        var benchmark = new MarkPflug65KCsvBenchmarks();
        benchmark.Setup();
        for (int index = 0; index < WarmupIterations; index++) {
            CsvReadObservation officeObservation = benchmark.OfficeIMO();
            CsvReadObservation excelReaderObservation = benchmark.ExcelReaderNet();
            ValidateEquivalent($"warmup {index}", officeObservation, excelReaderObservation);
        }

        var officeSamples = new double[iterations];
        var excelReaderSamples = new double[iterations];
        var pairedRatios = new double[iterations];
        for (int index = 0; index < iterations; index++) {
            (double Milliseconds, CsvReadObservation Observation) officeFirst;
            (double Milliseconds, CsvReadObservation Observation) officeSecond;
            (double Milliseconds, CsvReadObservation Observation) excelReaderFirst;
            (double Milliseconds, CsvReadObservation Observation) excelReaderSecond;
            if ((index & 1) == 0) {
                officeFirst = Measure(benchmark.OfficeIMO);
                excelReaderFirst = Measure(benchmark.ExcelReaderNet);
                excelReaderSecond = Measure(benchmark.ExcelReaderNet);
                officeSecond = Measure(benchmark.OfficeIMO);
            } else {
                excelReaderFirst = Measure(benchmark.ExcelReaderNet);
                officeFirst = Measure(benchmark.OfficeIMO);
                officeSecond = Measure(benchmark.OfficeIMO);
                excelReaderSecond = Measure(benchmark.ExcelReaderNet);
            }

            ValidateEquivalent($"sample {index} first", officeFirst.Observation, excelReaderFirst.Observation);
            ValidateEquivalent($"sample {index} second", officeSecond.Observation, excelReaderSecond.Observation);
            officeSamples[index] = (officeFirst.Milliseconds + officeSecond.Milliseconds) / 2d;
            excelReaderSamples[index] = (excelReaderFirst.Milliseconds + excelReaderSecond.Milliseconds) / 2d;
            pairedRatios[index] = officeSamples[index] / excelReaderSamples[index];
        }

        double officeMedian = Median(officeSamples);
        double excelReaderMedian = Median(excelReaderSamples);
        Console.WriteLine(FormattableString.Invariant(
            $"Paired CSV comparison ({WarmupIterations} warmups, {iterations} ABBA samples, affinity {affinity}, priority {priority}): OfficeIMO median {officeMedian:F3} ms, ExcelReader.NET median {excelReaderMedian:F3} ms, ratio of medians {officeMedian / excelReaderMedian:F4}, paired ratio median {Median(pairedRatios):F4} (P25 {Percentile(pairedRatios, 0.25d):F4}, P75 {Percentile(pairedRatios, 0.75d):F4})."));
    }

    private static (double Milliseconds, CsvReadObservation Observation) Measure(
        Func<CsvReadObservation> operation) {
        long started = Stopwatch.GetTimestamp();
        CsvReadObservation observation = operation();
        return (Stopwatch.GetElapsedTime(started).TotalMilliseconds, observation);
    }

    private static void ValidateEquivalent(
        string sample,
        CsvReadObservation officeObservation,
        CsvReadObservation excelReaderObservation) {
        if (officeObservation != excelReaderObservation) {
            throw new InvalidDataException(
                $"Paired CSV {sample} produced different observations: " +
                $"OfficeIMO={officeObservation}; ExcelReader.NET={excelReaderObservation}.");
        }
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
