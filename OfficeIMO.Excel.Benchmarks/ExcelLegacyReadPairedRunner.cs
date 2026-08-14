using System.Diagnostics;
using System.Globalization;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

internal static class ExcelLegacyReadPairedRunner {
    private const int WarmupIterations = 12;

    internal static void Run(string[] arguments, bool useXlsb) {
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

        Func<ExcelReadObservation> runOfficeIMO;
        Func<ExcelReadObservation> runSylvan;
        string format;
        if (useXlsb) {
            var benchmark = new MarkPflug65KXlsbBenchmarks();
            benchmark.Setup();
            runOfficeIMO = benchmark.OfficeIMO;
            runSylvan = benchmark.Sylvan;
            format = "XLSB";
        } else {
            var benchmark = new MarkPflug65KXlsBenchmarks();
            benchmark.Setup();
            runOfficeIMO = benchmark.OfficeIMO;
            runSylvan = benchmark.Sylvan;
            format = "XLS";
        }

        for (int index = 0; index < WarmupIterations; index++) {
            ExcelReadObservation officeObservation = runOfficeIMO();
            ExcelReadObservation sylvanObservation = runSylvan();
            ValidateEquivalent(format, $"warmup {index}", officeObservation, sylvanObservation);
        }

        var officeSamples = new double[iterations];
        var sylvanSamples = new double[iterations];
        var pairedRatios = new double[iterations];
        for (int index = 0; index < iterations; index++) {
            (double Milliseconds, ExcelReadObservation Observation) officeFirst;
            (double Milliseconds, ExcelReadObservation Observation) officeSecond;
            (double Milliseconds, ExcelReadObservation Observation) sylvanFirst;
            (double Milliseconds, ExcelReadObservation Observation) sylvanSecond;
            if ((index & 1) == 0) {
                officeFirst = Measure(runOfficeIMO);
                sylvanFirst = Measure(runSylvan);
                sylvanSecond = Measure(runSylvan);
                officeSecond = Measure(runOfficeIMO);
            } else {
                sylvanFirst = Measure(runSylvan);
                officeFirst = Measure(runOfficeIMO);
                officeSecond = Measure(runOfficeIMO);
                sylvanSecond = Measure(runSylvan);
            }

            ValidateEquivalent(format, $"sample {index} first", officeFirst.Observation, sylvanFirst.Observation);
            ValidateEquivalent(format, $"sample {index} second", officeSecond.Observation, sylvanSecond.Observation);
            officeSamples[index] = (officeFirst.Milliseconds + officeSecond.Milliseconds) / 2d;
            sylvanSamples[index] = (sylvanFirst.Milliseconds + sylvanSecond.Milliseconds) / 2d;
            pairedRatios[index] = officeSamples[index] / sylvanSamples[index];
        }

        double officeMedian = Median(officeSamples);
        double sylvanMedian = Median(sylvanSamples);
        Console.WriteLine(string.Format(
            CultureInfo.InvariantCulture,
            "Paired {0} comparison ({1} warmups, {2} ABBA samples, affinity {3}, priority {4}): " +
            "OfficeIMO median {5:F3} ms, Sylvan median {6:F3} ms, ratio of medians {7:F4}, " +
            "paired ratio median {8:F4} (P25 {9:F4}, P75 {10:F4}).",
            format,
            WarmupIterations,
            iterations,
            affinity,
            priority,
            officeMedian,
            sylvanMedian,
            officeMedian / sylvanMedian,
            Median(pairedRatios),
            Percentile(pairedRatios, 0.25d),
            Percentile(pairedRatios, 0.75d)));
    }

    private static (double Milliseconds, ExcelReadObservation Observation) Measure(
        Func<ExcelReadObservation> operation) {
        long started = Stopwatch.GetTimestamp();
        ExcelReadObservation observation = operation();
        return (Stopwatch.GetElapsedTime(started).TotalMilliseconds, observation);
    }

    private static void ValidateEquivalent(
        string format,
        string sample,
        ExcelReadObservation officeObservation,
        ExcelReadObservation sylvanObservation) {
        if (officeObservation != sylvanObservation) {
            throw new InvalidDataException(
                $"Paired {format} {sample} produced different observations: " +
                $"OfficeIMO={officeObservation}; Sylvan={sylvanObservation}.");
        }
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
