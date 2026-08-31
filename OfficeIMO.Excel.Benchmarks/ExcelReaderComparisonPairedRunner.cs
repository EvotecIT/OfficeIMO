using System.Diagnostics;
using System.Globalization;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelReaderComparisonPairedRunner {
    private const int WarmupIterations = 16;

    internal static void Run(string[] arguments, ExcelFileFormat format) {
        int iterations = arguments.Length > 1 && int.TryParse(arguments[1], out int parsedIterations)
            ? parsedIterations
            : 40;
        if (iterations <= 0) {
            throw new ArgumentOutOfRangeException(nameof(iterations));
        }

        bool prefetch = arguments.Length > 4
            && string.Equals(arguments[4], "prefetch", StringComparison.OrdinalIgnoreCase);
        if (prefetch && format != ExcelFileFormat.Xlsx) {
            throw new ArgumentException(
                "Worksheet prefetch comparison is supported only for XLSX.",
                nameof(arguments));
        }

        string affinity = arguments.Length > 2
            ? BenchmarkProcessorAffinity.Apply(arguments[2])
            : "unchanged";
        string priority = arguments.Length > 3
            && !string.Equals(arguments[3], "unchanged", StringComparison.OrdinalIgnoreCase)
                ? BenchmarkProcessorAffinity.ApplyPriority(arguments[3])
                : Process.GetCurrentProcess().PriorityClass.ToString();
        (Func<ExcelReadObservation> RunOfficeIMO, Func<ExcelReadObservation> RunExcelReader) =
            CreateOperations(format, prefetch);

        for (int index = 0; index < WarmupIterations; index++) {
            ExcelReadObservation officeObservation = RunOfficeIMO();
            ExcelReadObservation excelReaderObservation = RunExcelReader();
            ValidateEquivalent(format, $"warmup {index}", officeObservation, excelReaderObservation);
        }

        var officeSamples = new double[iterations];
        var excelReaderSamples = new double[iterations];
        var pairedRatios = new double[iterations];
        for (int index = 0; index < iterations; index++) {
            (double Milliseconds, ExcelReadObservation Observation) officeFirst;
            (double Milliseconds, ExcelReadObservation Observation) officeSecond;
            (double Milliseconds, ExcelReadObservation Observation) excelReaderFirst;
            (double Milliseconds, ExcelReadObservation Observation) excelReaderSecond;
            if ((index & 1) == 0) {
                officeFirst = Measure(RunOfficeIMO);
                excelReaderFirst = Measure(RunExcelReader);
                excelReaderSecond = Measure(RunExcelReader);
                officeSecond = Measure(RunOfficeIMO);
            } else {
                excelReaderFirst = Measure(RunExcelReader);
                officeFirst = Measure(RunOfficeIMO);
                officeSecond = Measure(RunOfficeIMO);
                excelReaderSecond = Measure(RunExcelReader);
            }

            ValidateEquivalent(format, $"sample {index} first", officeFirst.Observation, excelReaderFirst.Observation);
            ValidateEquivalent(format, $"sample {index} second", officeSecond.Observation, excelReaderSecond.Observation);
            officeSamples[index] = (officeFirst.Milliseconds + officeSecond.Milliseconds) / 2d;
            excelReaderSamples[index] = (excelReaderFirst.Milliseconds + excelReaderSecond.Milliseconds) / 2d;
            pairedRatios[index] = officeSamples[index] / excelReaderSamples[index];
        }

        double officeMedian = Median(officeSamples);
        double excelReaderMedian = Median(excelReaderSamples);
        Console.WriteLine(string.Format(
            CultureInfo.InvariantCulture,
            "Paired {0} comparison{11} ({1} warmups, {2} ABBA samples, affinity {3}, priority {4}): " +
            "OfficeIMO median {5:F3} ms, ExcelReader.NET median {6:F3} ms, ratio of medians {7:F4}, " +
            "paired ratio median {8:F4} (P25 {9:F4}, P75 {10:F4}).",
            format.ToString().ToUpperInvariant(),
            WarmupIterations,
            iterations,
            affinity,
            priority,
            officeMedian,
            excelReaderMedian,
            officeMedian / excelReaderMedian,
            Median(pairedRatios),
            Percentile(pairedRatios, 0.25d),
            Percentile(pairedRatios, 0.75d),
            prefetch ? " with bounded prefetch enabled for both libraries" : string.Empty));
    }

    private static (Func<ExcelReadObservation> RunOfficeIMO, Func<ExcelReadObservation> RunExcelReader)
        CreateOperations(ExcelFileFormat format, bool prefetch) {
        switch (format) {
            case ExcelFileFormat.Xlsx: {
                var benchmark = new MarkPflug65KXlsxBenchmarks();
                benchmark.Setup();
                return prefetch
                    ? (benchmark.OfficeIMO_Prefetch, benchmark.ExcelReaderNet_Prefetch)
                    : (benchmark.OfficeIMO, benchmark.ExcelReaderNet);
            }
            case ExcelFileFormat.Xlsb: {
                var benchmark = new MarkPflug65KXlsbBenchmarks();
                benchmark.Setup();
                return (benchmark.OfficeIMO, benchmark.ExcelReaderNet);
            }
            case ExcelFileFormat.Xls: {
                var benchmark = new MarkPflug65KXlsBenchmarks();
                benchmark.Setup();
                return (benchmark.OfficeIMO, benchmark.ExcelReaderNet);
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(format), format, "Only XLSX, XLSB, and XLS are supported.");
        }
    }

    private static (double Milliseconds, ExcelReadObservation Observation) Measure(
        Func<ExcelReadObservation> operation) {
        long started = Stopwatch.GetTimestamp();
        ExcelReadObservation observation = operation();
        return (Stopwatch.GetElapsedTime(started).TotalMilliseconds, observation);
    }

    private static void ValidateEquivalent(
        ExcelFileFormat format,
        string sample,
        ExcelReadObservation officeObservation,
        ExcelReadObservation excelReaderObservation) {
        if (officeObservation != excelReaderObservation) {
            throw new InvalidDataException(
                $"Paired {format} {sample} produced different observations: " +
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
