using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;
using OfficeIMO.Benchmarks;
using OfficeIMO.CSV.Benchmarks;
using System.Runtime.Intrinsics.X86;

if (args.Length > 0
    && string.Equals(args[0], "--print-intrinsics", StringComparison.OrdinalIgnoreCase)) {
    Console.WriteLine($"AVX512BW={Avx512BW.IsSupported}; AVX2={Avx2.IsSupported}");
    return;
}

bool profileOfficeIMO = args.Length > 0 &&
    string.Equals(args[0], "--profile-markpflug65k-officeimo", StringComparison.OrdinalIgnoreCase);
bool profileSep = args.Length > 0 &&
    string.Equals(args[0], "--profile-markpflug65k-sep", StringComparison.OrdinalIgnoreCase);
bool profileSylvan = args.Length > 0 &&
    string.Equals(args[0], "--profile-markpflug65k-sylvan", StringComparison.OrdinalIgnoreCase);
bool comparePaired = args.Length > 0 &&
    string.Equals(args[0], "--compare-markpflug65k-paired", StringComparison.OrdinalIgnoreCase);

if (comparePaired) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 100;
    if (iterations <= 0) {
        throw new ArgumentOutOfRangeException(nameof(iterations));
    }
    ApplyProcessAffinity(args, argumentIndex: 2);
    const int warmupIterations = 10;
    string affinity = args.Length > 2 ? args[2] : "unchanged";

    var benchmark = new MarkPflug65KCsvBenchmarks();
    benchmark.Setup();
    for (int index = 0; index < warmupIterations; index++) {
        benchmark.OfficeIMO();
        benchmark.Sep();
        benchmark.Sylvan();
    }

    var officeSamples = new double[iterations];
    var sepSamples = new double[iterations];
    var sylvanSamples = new double[iterations];
    var officeSepRatios = new double[iterations];
    var officeSylvanRatios = new double[iterations];
    for (int index = 0; index < iterations; index++) {
        CsvReadObservation officeObservation;
        CsvReadObservation sepObservation;
        CsvReadObservation sylvanObservation;
        switch (index % 3) {
            case 0:
                officeSamples[index] = MeasureMilliseconds(benchmark.OfficeIMO, out officeObservation);
                sepSamples[index] = MeasureMilliseconds(benchmark.Sep, out sepObservation);
                sylvanSamples[index] = MeasureMilliseconds(benchmark.Sylvan, out sylvanObservation);
                break;
            case 1:
                sepSamples[index] = MeasureMilliseconds(benchmark.Sep, out sepObservation);
                sylvanSamples[index] = MeasureMilliseconds(benchmark.Sylvan, out sylvanObservation);
                officeSamples[index] = MeasureMilliseconds(benchmark.OfficeIMO, out officeObservation);
                break;
            default:
                sylvanSamples[index] = MeasureMilliseconds(benchmark.Sylvan, out sylvanObservation);
                officeSamples[index] = MeasureMilliseconds(benchmark.OfficeIMO, out officeObservation);
                sepSamples[index] = MeasureMilliseconds(benchmark.Sep, out sepObservation);
                break;
        }

        if (officeObservation != sepObservation || officeObservation != sylvanObservation) {
            throw new InvalidDataException(
                $"Paired CSV sample {index} produced different observations: OfficeIMO={officeObservation}; Sep={sepObservation}; Sylvan={sylvanObservation}.");
        }
        officeSepRatios[index] = officeSamples[index] / sepSamples[index];
        officeSylvanRatios[index] = officeSamples[index] / sylvanSamples[index];
    }

    double officeMedian = Median(officeSamples);
    double sepMedian = Median(sepSamples);
    double sylvanMedian = Median(sylvanSamples);
    Console.WriteLine(
        $"Paired CSV comparison ({warmupIterations} warmups, {iterations} rotating samples, affinity {affinity}, " +
        $"AVX512BW={Avx512BW.IsSupported}, AVX2={Avx2.IsSupported}): " +
        $"OfficeIMO {officeMedian:F3} ms, Sep {sepMedian:F3} ms, Sylvan {sylvanMedian:F3} ms; " +
        $"OfficeIMO/Sep paired median {Median(officeSepRatios):F4} " +
        $"(P25 {Percentile(officeSepRatios, 0.25d):F4}, P75 {Percentile(officeSepRatios, 0.75d):F4}); " +
        $"OfficeIMO/Sylvan paired median {Median(officeSylvanRatios):F4} " +
        $"(P25 {Percentile(officeSylvanRatios, 0.25d):F4}, P75 {Percentile(officeSylvanRatios, 0.75d):F4}).");
    return;
}

if (profileOfficeIMO || profileSep || profileSylvan) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 100;
    if (iterations <= 0) {
        throw new ArgumentOutOfRangeException(nameof(iterations));
    }
    ApplyProcessAffinity(args, argumentIndex: 2);

    var benchmark = new MarkPflug65KCsvBenchmarks();
    benchmark.Setup();
    Func<CsvReadObservation> run = profileOfficeIMO
        ? benchmark.OfficeIMO
        : profileSep
            ? benchmark.Sep
            : benchmark.Sylvan;
    string implementation = profileOfficeIMO
        ? "OfficeIMO"
        : profileSep
            ? "Sep"
            : "Sylvan";
    for (int index = 0; index < 3; index++) {
        run();
    }

    CsvReadObservation observation = default;
    var stopwatch = System.Diagnostics.Stopwatch.StartNew();
    for (int index = 0; index < iterations; index++) {
        observation = run();
    }
    stopwatch.Stop();

    Console.WriteLine(
        $"Profiled {implementation} CSV {iterations} times in {stopwatch.Elapsed.TotalMilliseconds:F2} ms " +
        $"({stopwatch.Elapsed.TotalMilliseconds / iterations:F3} ms/iteration): {observation}.");
    return;
}

var config = ManualConfig
    .Create(DefaultConfig.Instance)
    .AddDiagnoser(MemoryDiagnoser.Default)
    .AddExporter(JsonExporter.Full)
    .WithSummaryStyle(SummaryStyle.Default.WithRatioStyle(RatioStyle.Percentage))
    .AddColumn(StatisticColumn.OperationsPerSecond);

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args, config);

static double MeasureMilliseconds(
    Func<CsvReadObservation> operation,
    out CsvReadObservation observation) {
    long started = System.Diagnostics.Stopwatch.GetTimestamp();
    observation = operation();
    return System.Diagnostics.Stopwatch.GetElapsedTime(started).TotalMilliseconds;
}

static void ApplyProcessAffinity(string[] arguments, int argumentIndex) {
    if (arguments.Length <= argumentIndex
        || !long.TryParse(arguments[argumentIndex], out long affinityMask)) {
        return;
    }
    if (!OperatingSystem.IsWindows()) {
        throw new PlatformNotSupportedException(
            "Processor-affinity comparison is available only on Windows.");
    }
    if (affinityMask <= 0) {
        throw new ArgumentOutOfRangeException(nameof(affinityMask));
    }

    System.Diagnostics.Process.GetCurrentProcess().ProcessorAffinity = checked((nint)affinityMask);
}

static double Median(double[] samples) {
    Array.Sort(samples);
    int middle = samples.Length / 2;
    return (samples.Length & 1) == 0
        ? (samples[middle - 1] + samples[middle]) / 2d
        : samples[middle];
}

static double Percentile(double[] samples, double percentile) {
    if (samples.Length == 0) {
        throw new ArgumentException("At least one sample is required.", nameof(samples));
    }
    if (percentile < 0d || percentile > 1d) {
        throw new ArgumentOutOfRangeException(nameof(percentile));
    }

    Array.Sort(samples);
    double position = (samples.Length - 1) * percentile;
    int lower = (int)position;
    int upper = Math.Min(lower + 1, samples.Length - 1);
    double fraction = position - lower;
    return samples[lower] + (samples[upper] - samples[lower]) * fraction;
}
