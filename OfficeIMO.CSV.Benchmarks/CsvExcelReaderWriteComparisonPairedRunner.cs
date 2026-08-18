using System.Diagnostics;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.CSV.Benchmarks;

internal static class CsvExcelReaderWriteComparisonPairedRunner {
    private const int WarmupIterations = 12;

    internal static void Run(string[] arguments) {
        int rowCount = arguments.Length > 1 && int.TryParse(arguments[1], out int parsedRowCount)
            ? parsedRowCount
            : 25_000;
        CsvBenchmarkShape shape = arguments.Length > 2
            && Enum.TryParse(arguments[2], ignoreCase: true, out CsvBenchmarkShape parsedShape)
                ? parsedShape
                : CsvBenchmarkShape.Mixed;
        int iterations = arguments.Length > 3 && int.TryParse(arguments[3], out int parsedIterations)
            ? parsedIterations
            : 30;
        if (rowCount <= 0) {
            throw new ArgumentOutOfRangeException(nameof(rowCount));
        }
        if (iterations <= 0) {
            throw new ArgumentOutOfRangeException(nameof(iterations));
        }

        string affinity = arguments.Length > 4
            ? BenchmarkProcessorAffinity.Apply(arguments[4])
            : "unchanged";
        string priority = arguments.Length > 5
            && !string.Equals(arguments[5], "unchanged", StringComparison.OrdinalIgnoreCase)
                ? BenchmarkProcessorAffinity.ApplyPriority(arguments[5])
                : Process.GetCurrentProcess().PriorityClass.ToString();
        int invocationsPerLeg = arguments.Length > 6 && int.TryParse(arguments[6], out int parsedInvocations)
            ? parsedInvocations
            : 8;
        if (invocationsPerLeg <= 0) {
            throw new ArgumentOutOfRangeException(nameof(invocationsPerLeg));
        }

        var benchmark = new CsvBenchmarks { RowCount = rowCount, Shape = shape };
        benchmark.Setup();
        for (int index = 0; index < WarmupIterations; index++) {
            ValidateResult($"OfficeIMO warmup {index}", benchmark.OfficeIMO_WriteProjectedRowsUtf8());
            ValidateResult($"ExcelReader.NET warmup {index}", benchmark.ExcelReaderNet_WriteProjectedRows());
        }

        var officeSamples = new double[iterations];
        var excelReaderSamples = new double[iterations];
        var pairedRatios = new double[iterations];
        for (int index = 0; index < iterations; index++) {
            (double Milliseconds, int Result) officeFirst;
            (double Milliseconds, int Result) officeSecond;
            (double Milliseconds, int Result) excelReaderFirst;
            (double Milliseconds, int Result) excelReaderSecond;
            if ((index & 1) == 0) {
                officeFirst = Measure(benchmark.OfficeIMO_WriteProjectedRowsUtf8, invocationsPerLeg);
                excelReaderFirst = Measure(benchmark.ExcelReaderNet_WriteProjectedRows, invocationsPerLeg);
                excelReaderSecond = Measure(benchmark.ExcelReaderNet_WriteProjectedRows, invocationsPerLeg);
                officeSecond = Measure(benchmark.OfficeIMO_WriteProjectedRowsUtf8, invocationsPerLeg);
            } else {
                excelReaderFirst = Measure(benchmark.ExcelReaderNet_WriteProjectedRows, invocationsPerLeg);
                officeFirst = Measure(benchmark.OfficeIMO_WriteProjectedRowsUtf8, invocationsPerLeg);
                officeSecond = Measure(benchmark.OfficeIMO_WriteProjectedRowsUtf8, invocationsPerLeg);
                excelReaderSecond = Measure(benchmark.ExcelReaderNet_WriteProjectedRows, invocationsPerLeg);
            }

            ValidateResult($"OfficeIMO sample {index} first", officeFirst.Result);
            ValidateResult($"OfficeIMO sample {index} second", officeSecond.Result);
            ValidateResult($"ExcelReader.NET sample {index} first", excelReaderFirst.Result);
            ValidateResult($"ExcelReader.NET sample {index} second", excelReaderSecond.Result);
            officeSamples[index] = (officeFirst.Milliseconds + officeSecond.Milliseconds) / 2d;
            excelReaderSamples[index] = (excelReaderFirst.Milliseconds + excelReaderSecond.Milliseconds) / 2d;
            pairedRatios[index] = officeSamples[index] / excelReaderSamples[index];
        }

        double officeMedian = Median(officeSamples);
        double excelReaderMedian = Median(excelReaderSamples);
        Console.WriteLine(FormattableString.Invariant(
            $"Paired projected-row CSV write comparison ({shape}, {rowCount:N0} rows, {WarmupIterations} warmups, {iterations} ABBA samples, {invocationsPerLeg} invocations per leg, affinity {affinity}, priority {priority}): OfficeIMO median {officeMedian:F3} ms, ExcelReader.NET median {excelReaderMedian:F3} ms, ratio of medians {officeMedian / excelReaderMedian:F4}, paired ratio median {Median(pairedRatios):F4} (P25 {Percentile(pairedRatios, 0.25d):F4}, P75 {Percentile(pairedRatios, 0.75d):F4})."));
    }

    private static (double Milliseconds, int Result) Measure(Func<int> operation, int invocationCount) {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        long started = Stopwatch.GetTimestamp();
        int result = 0;
        for (int index = 0; index < invocationCount; index++) {
            result = operation();
        }
        return (Stopwatch.GetElapsedTime(started).TotalMilliseconds / invocationCount, result);
    }

    private static void ValidateResult(string sample, int result) {
        if (result <= 0) {
            throw new InvalidDataException($"CSV {sample} produced an invalid result of {result}.");
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
