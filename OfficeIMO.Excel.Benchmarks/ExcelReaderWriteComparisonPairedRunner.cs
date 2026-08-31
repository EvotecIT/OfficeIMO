using System.Diagnostics;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelReaderWriteComparisonPairedRunner {
    private const int WarmupIterations = 12;

    internal static void Run(string[] arguments) {
        ExcelFileFormat format = arguments.Length > 1
            && Enum.TryParse(arguments[1], ignoreCase: true, out ExcelFileFormat parsedFormat)
                ? parsedFormat
                : ExcelFileFormat.Xlsx;
        if (format is not (ExcelFileFormat.Xlsx or ExcelFileFormat.Xlsb or ExcelFileFormat.Xls)) {
            throw new ArgumentOutOfRangeException(nameof(format), format, "Only XLSX, XLSB, and XLS are supported.");
        }

        int rowCount = arguments.Length > 2 && int.TryParse(arguments[2], out int parsedRowCount)
            ? parsedRowCount
            : 25_000;
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

        ComparisonOperations operations = CreateOperations(format, rowCount);
        if (operations.Conformance is { IsEquivalent: false } nonEquivalent) {
            Console.WriteLine(
                $"ExcelReader.NET {format.ToString().ToUpperInvariant()} conformance probe: " +
                $"semantic={nonEquivalent.SemanticRoundTrip}, structural={nonEquivalent.StructurallyConformant}. " +
                $"{nonEquivalent.Detail} Artifact bytes: OfficeIMO={nonEquivalent.OfficeOutputBytes:N0}, " +
                $"ExcelReader.NET={nonEquivalent.CompetitorOutputBytes:N0}.");
            Console.WriteLine(
                "Paired timing withheld because the generated workbooks are not equivalent. " +
                "This conformance probe remains active and will automatically enter the timed lane when a future competitor release produces an equivalent workbook.");
            if (operations.ProfileOfficeIMO is not null) {
                Console.WriteLine(
                    "OfficeIMO one-pass stage profile (diagnostic, outside comparison timing): " +
                    string.Join(", ", operations.ProfileOfficeIMO()
                        .Select(static stage => FormattableString.Invariant($"{stage.Name}={stage.Milliseconds:F3} ms"))));
            }
            return;
        }

        (Func<int> RunOfficeIMO, Func<int> RunExcelReader) = operations;
        for (int index = 0; index < WarmupIterations; index++) {
            ValidateResult(format, rowCount, $"OfficeIMO warmup {index}", RunOfficeIMO());
            ValidateResult(format, rowCount, $"ExcelReader.NET warmup {index}", RunExcelReader());
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
                officeFirst = Measure(RunOfficeIMO, invocationsPerLeg);
                excelReaderFirst = Measure(RunExcelReader, invocationsPerLeg);
                excelReaderSecond = Measure(RunExcelReader, invocationsPerLeg);
                officeSecond = Measure(RunOfficeIMO, invocationsPerLeg);
            } else {
                excelReaderFirst = Measure(RunExcelReader, invocationsPerLeg);
                officeFirst = Measure(RunOfficeIMO, invocationsPerLeg);
                officeSecond = Measure(RunOfficeIMO, invocationsPerLeg);
                excelReaderSecond = Measure(RunExcelReader, invocationsPerLeg);
            }

            ValidateResult(format, rowCount, $"OfficeIMO sample {index} first", officeFirst.Result);
            ValidateResult(format, rowCount, $"OfficeIMO sample {index} second", officeSecond.Result);
            ValidateResult(format, rowCount, $"ExcelReader.NET sample {index} first", excelReaderFirst.Result);
            ValidateResult(format, rowCount, $"ExcelReader.NET sample {index} second", excelReaderSecond.Result);
            officeSamples[index] = (officeFirst.Milliseconds + officeSecond.Milliseconds) / 2d;
            excelReaderSamples[index] = (excelReaderFirst.Milliseconds + excelReaderSecond.Milliseconds) / 2d;
            pairedRatios[index] = officeSamples[index] / excelReaderSamples[index];
        }

        double officeMedian = Median(officeSamples);
        double excelReaderMedian = Median(excelReaderSamples);
        Console.WriteLine(FormattableString.Invariant(
            $"Paired {format.ToString().ToUpperInvariant()} write comparison, validated equivalent lane ({rowCount:N0} rows, {WarmupIterations} warmups, {iterations} ABBA samples, {invocationsPerLeg} invocations per leg, affinity {affinity}, priority {priority}): OfficeIMO median {officeMedian:F3} ms, ExcelReader.NET median {excelReaderMedian:F3} ms, ratio of medians {officeMedian / excelReaderMedian:F4}, paired ratio median {Median(pairedRatios):F4} (P25 {Percentile(pairedRatios, 0.25d):F4}, P75 {Percentile(pairedRatios, 0.75d):F4})."));
        if (operations.Conformance is { } conformance) {
            Console.WriteLine(
                $"ExcelReader.NET conformance: semantic={conformance.SemanticRoundTrip}, " +
                $"structural={conformance.StructurallyConformant}. {conformance.Detail} " +
                $"Artifact bytes: OfficeIMO={conformance.OfficeOutputBytes:N0}, " +
                $"ExcelReader.NET={conformance.CompetitorOutputBytes:N0}.");
        }
        if (operations.ProfileOfficeIMO is not null) {
            Console.WriteLine(
                "OfficeIMO one-pass stage profile (diagnostic, outside timed samples): " +
                string.Join(", ", operations.ProfileOfficeIMO()
                    .Select(static stage => FormattableString.Invariant($"{stage.Name}={stage.Milliseconds:F3} ms"))));
        }
    }

    private static ComparisonOperations CreateOperations(
        ExcelFileFormat format,
        int rowCount) {
        if (format == ExcelFileFormat.Xlsx) {
            var benchmark = new ExcelGeneratedRowStreamingBenchmarks { RowCount = rowCount };
            benchmark.Setup();
            return new ComparisonOperations(
                benchmark.WriteRowsGenerated,
                benchmark.ExcelReaderNetWriteRowsGenerated,
                Conformance: null,
                ProfileOfficeIMO: null);
        }

        var binaryBenchmark = new ExcelNativeBinaryWriteBenchmarks {
            Format = format,
            RowCount = rowCount
        };
        BinaryWriteConformanceObservation conformance = binaryBenchmark.SetupComparison();
        return new ComparisonOperations(
            binaryBenchmark.OfficeIMO_PublicTabularWrite,
            binaryBenchmark.ExcelReaderNet_DiagnosticWrite,
            conformance,
            binaryBenchmark.ProfileOfficeIMOWriteStages);
    }

    private static (double Milliseconds, int Result) Measure(Func<int> operation, int invocationCount) {
        BenchmarkMeasurement.PrepareForMeasurement();
        long started = Stopwatch.GetTimestamp();
        int result = 0;
        for (int index = 0; index < invocationCount; index++) {
            result = operation();
        }
        return (Stopwatch.GetElapsedTime(started).TotalMilliseconds / invocationCount, result);
    }

    private static void ValidateResult(ExcelFileFormat format, int expectedRowCount, string sample, int result) {
        if (format == ExcelFileFormat.Xlsx && result != expectedRowCount) {
            throw new InvalidDataException(
                $"{format} {sample} reported {result} rows instead of {expectedRowCount}.");
        }

        if (result <= 0) {
            throw new InvalidDataException($"{format} {sample} produced an invalid result of {result}.");
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

    private readonly record struct ComparisonOperations(
        Func<int> RunOfficeIMO,
        Func<int> RunExcelReader,
        BinaryWriteConformanceObservation? Conformance,
        Func<IReadOnlyList<(string Name, double Milliseconds)>>? ProfileOfficeIMO) {
        internal void Deconstruct(out Func<int> officeIMO, out Func<int> excelReader) {
            officeIMO = RunOfficeIMO;
            excelReader = RunExcelReader;
        }
    }
}
