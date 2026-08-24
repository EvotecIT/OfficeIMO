using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;
using OfficeIMO.Benchmarks;
using OfficeIMO.CSV.Benchmarks;
using System.Runtime.Intrinsics.X86;

if (args.Length > 0 && string.Equals(args[0], "--datareader-write-size-evidence", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = CsvDataReaderWriteSizeEvidenceRunner.Run(args.Skip(1).ToArray());
    return;
}

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
bool profileTrimOfficeIMO = args.Length > 0 &&
    string.Equals(args[0], "--profile-trim-unescape-strings-officeimo", StringComparison.OrdinalIgnoreCase);
bool profileTrimSep = args.Length > 0 &&
    string.Equals(args[0], "--profile-trim-unescape-strings-sep", StringComparison.OrdinalIgnoreCase);
bool profileTrimSpanOfficeIMO = args.Length > 0 &&
    string.Equals(args[0], "--profile-trim-unescape-spans-officeimo", StringComparison.OrdinalIgnoreCase);
bool profileTrimSpanSep = args.Length > 0 &&
    string.Equals(args[0], "--profile-trim-unescape-spans-sep", StringComparison.OrdinalIgnoreCase);
bool profileTypedOfficeRowsAs = args.Length > 0 &&
    string.Equals(args[0], "--profile-typed-officeimo-rowsas", StringComparison.OrdinalIgnoreCase);
bool profileTypedOfficeManual = args.Length > 0 &&
    string.Equals(args[0], "--profile-typed-officeimo-manual", StringComparison.OrdinalIgnoreCase);
bool profileTypedOfficeParallel = args.Length > 0 &&
    string.Equals(args[0], "--profile-typed-officeimo-parallel", StringComparison.OrdinalIgnoreCase);
bool profileTypedSepSequential = args.Length > 0 &&
    string.Equals(args[0], "--profile-typed-sep-sequential", StringComparison.OrdinalIgnoreCase);
bool profileTypedSepParallel = args.Length > 0 &&
    string.Equals(args[0], "--profile-typed-sep-parallel", StringComparison.OrdinalIgnoreCase);
bool comparePaired = args.Length > 0 &&
    string.Equals(args[0], "--compare-markpflug65k-paired", StringComparison.OrdinalIgnoreCase);
bool compareExcelReaderPaired = args.Length > 0 &&
    string.Equals(args[0], "--compare-excelreader-paired", StringComparison.OrdinalIgnoreCase);
bool compareExcelReaderWritePaired = args.Length > 0 &&
    string.Equals(args[0], "--compare-excelreader-write-paired", StringComparison.OrdinalIgnoreCase);
bool compareTrimUnescapeStringsPaired = args.Length > 0 &&
    string.Equals(args[0], "--compare-trim-unescape-strings-paired", StringComparison.OrdinalIgnoreCase);
bool compareTypedSequentialPaired = args.Length > 0 &&
    string.Equals(args[0], "--compare-typed-sequential-paired", StringComparison.OrdinalIgnoreCase);
bool compareTypedParallelPaired = args.Length > 0 &&
    string.Equals(args[0], "--compare-typed-parallel-paired", StringComparison.OrdinalIgnoreCase);
bool compareDataReaderWritePaired = args.Length > 0 &&
    string.Equals(args[0], "--compare-datareader-write-paired", StringComparison.OrdinalIgnoreCase);
bool compareDataReaderParallelWritePaired = args.Length > 0 &&
    string.Equals(args[0], "--compare-datareader-parallel-write-paired", StringComparison.OrdinalIgnoreCase);

if (compareExcelReaderPaired) {
    CsvExcelReaderComparisonPairedRunner.Run(args);
    return;
}

if (compareExcelReaderWritePaired) {
    CsvExcelReaderWriteComparisonPairedRunner.Run(args);
    return;
}

if (compareDataReaderParallelWritePaired) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 30;
    string affinity = ApplyProcessAffinity(args, argumentIndex: 2);
    string priority = ApplyProcessPriority(args, argumentIndex: 3);
    int rowCount = args.Length > 4 && int.TryParse(args[4], out int parsedRowCount)
        ? parsedRowCount
        : 100_000;
    int parallelDegree = args.Length > 5 && int.TryParse(args[5], out int parsedDegree)
        ? parsedDegree
        : 4;
    int parallelBatchSize = args.Length > 6 && int.TryParse(args[6], out int parsedBatchSize)
        ? parsedBatchSize
        : 4096;
    int invocationsPerLeg = args.Length > 7 && int.TryParse(args[7], out int parsedInvocations)
        ? parsedInvocations
        : 4;
    if (iterations <= 0) throw new ArgumentOutOfRangeException(nameof(iterations));
    if (rowCount <= 0) throw new ArgumentOutOfRangeException(nameof(rowCount));
    if (parallelDegree <= 0) throw new ArgumentOutOfRangeException(nameof(parallelDegree));
    if (parallelBatchSize <= 0) throw new ArgumentOutOfRangeException(nameof(parallelBatchSize));
    if (invocationsPerLeg <= 0) throw new ArgumentOutOfRangeException(nameof(invocationsPerLeg));
    const int warmupIterations = 8;

    using var measuredProcess = System.Diagnostics.Process.GetCurrentProcess();
    foreach (CsvBenchmarkShape shape in Enum.GetValues<CsvBenchmarkShape>()) {
        var benchmark = new CsvDataReaderWriteBenchmarks {
            RowCount = rowCount,
            Shape = shape,
            ParallelDegree = parallelDegree,
            ParallelBatchSize = parallelBatchSize
        };
        benchmark.SetupOfficeIMOSequentialAndParallel();
        for (int index = 0; index < warmupIterations; index++) {
            benchmark.OfficeIMO_WriteDataReader();
            benchmark.OfficeIMO_WriteDataReaderParallel();
        }

        var sequentialSamples = new double[iterations];
        var parallelSamples = new double[iterations];
        var wallRatios = new double[iterations];
        var sequentialCpuSamples = new double[iterations];
        var parallelCpuSamples = new double[iterations];
        var cpuRatios = new double[iterations];
        for (int index = 0; index < iterations; index++) {
            (double WallMilliseconds, double CpuMilliseconds) sequentialFirst;
            (double WallMilliseconds, double CpuMilliseconds) sequentialSecond;
            (double WallMilliseconds, double CpuMilliseconds) parallelFirst;
            (double WallMilliseconds, double CpuMilliseconds) parallelSecond;
            int sequentialResult;
            int parallelResult;
            if ((index & 1) == 0) {
                sequentialFirst = MeasureBatchValue(benchmark.OfficeIMO_WriteDataReader, invocationsPerLeg, measuredProcess, out sequentialResult);
                parallelFirst = MeasureBatchValue(benchmark.OfficeIMO_WriteDataReaderParallel, invocationsPerLeg, measuredProcess, out parallelResult);
                parallelSecond = MeasureBatchValue(benchmark.OfficeIMO_WriteDataReaderParallel, invocationsPerLeg, measuredProcess, out parallelResult);
                sequentialSecond = MeasureBatchValue(benchmark.OfficeIMO_WriteDataReader, invocationsPerLeg, measuredProcess, out sequentialResult);
            } else {
                parallelFirst = MeasureBatchValue(benchmark.OfficeIMO_WriteDataReaderParallel, invocationsPerLeg, measuredProcess, out parallelResult);
                sequentialFirst = MeasureBatchValue(benchmark.OfficeIMO_WriteDataReader, invocationsPerLeg, measuredProcess, out sequentialResult);
                sequentialSecond = MeasureBatchValue(benchmark.OfficeIMO_WriteDataReader, invocationsPerLeg, measuredProcess, out sequentialResult);
                parallelSecond = MeasureBatchValue(benchmark.OfficeIMO_WriteDataReaderParallel, invocationsPerLeg, measuredProcess, out parallelResult);
            }
            if (sequentialResult != parallelResult) {
                throw new InvalidDataException($"Paired DataReader write sample {index} produced different lengths: sequential={sequentialResult}; parallel={parallelResult}.");
            }

            sequentialSamples[index] = (sequentialFirst.WallMilliseconds + sequentialSecond.WallMilliseconds) / 2d;
            parallelSamples[index] = (parallelFirst.WallMilliseconds + parallelSecond.WallMilliseconds) / 2d;
            sequentialCpuSamples[index] = (sequentialFirst.CpuMilliseconds + sequentialSecond.CpuMilliseconds) / 2d;
            parallelCpuSamples[index] = (parallelFirst.CpuMilliseconds + parallelSecond.CpuMilliseconds) / 2d;
            wallRatios[index] = parallelSamples[index] / sequentialSamples[index];
            cpuRatios[index] = parallelCpuSamples[index] / sequentialCpuSamples[index];
        }

        Console.WriteLine(
            $"Paired CSV DataReader sequential/parallel write ({shape}, {rowCount} rows, {warmupIterations} warmups, {iterations} ABBA samples, {invocationsPerLeg} invocations per leg, affinity {affinity}, priority {priority}, DOP {parallelDegree}, batch {parallelBatchSize}): " +
            $"sequential wall median {Median(sequentialSamples):F3} ms, parallel wall median {Median(parallelSamples):F3} ms, " +
            $"parallel/sequential paired wall ratio {Median(wallRatios):F4} (P25 {Percentile(wallRatios, 0.25d):F4}, P75 {Percentile(wallRatios, 0.75d):F4}); " +
            $"sequential CPU median {Median(sequentialCpuSamples):F3} ms, parallel CPU median {Median(parallelCpuSamples):F3} ms, " +
            $"CPU ratio {Median(cpuRatios):F4} (P25 {Percentile(cpuRatios, 0.25d):F4}, P75 {Percentile(cpuRatios, 0.75d):F4}).");
    }
    return;
}

if (compareDataReaderWritePaired) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 40;
    if (iterations <= 0) throw new ArgumentOutOfRangeException(nameof(iterations));
    string affinity = ApplyProcessAffinity(args, argumentIndex: 2);
    string priority = ApplyProcessPriority(args, argumentIndex: 3);
    const int warmupIterations = 12;
    const int invocationsPerLeg = 8;

    foreach (CsvBenchmarkShape shape in Enum.GetValues<CsvBenchmarkShape>()) {
        var benchmark = new CsvDataReaderWriteBenchmarks { RowCount = 25_000, Shape = shape };
        benchmark.SetupOfficeIMOAndSylvan();
        for (int index = 0; index < warmupIterations; index++) {
            benchmark.OfficeIMO_WriteDataReader();
            benchmark.Sylvan_WriteDataReader();
        }

        var officeSamples = new double[iterations];
        var sylvanSamples = new double[iterations];
        var pairedRatios = new double[iterations];
        for (int index = 0; index < iterations; index++) {
            double officeFirst;
            double officeSecond;
            double sylvanFirst;
            double sylvanSecond;
            if ((index & 1) == 0) {
                officeFirst = MeasureMillisecondsBatchValue(benchmark.OfficeIMO_WriteDataReader, invocationsPerLeg, out _);
                sylvanFirst = MeasureMillisecondsBatchValue(benchmark.Sylvan_WriteDataReader, invocationsPerLeg, out _);
                sylvanSecond = MeasureMillisecondsBatchValue(benchmark.Sylvan_WriteDataReader, invocationsPerLeg, out _);
                officeSecond = MeasureMillisecondsBatchValue(benchmark.OfficeIMO_WriteDataReader, invocationsPerLeg, out _);
            } else {
                sylvanFirst = MeasureMillisecondsBatchValue(benchmark.Sylvan_WriteDataReader, invocationsPerLeg, out _);
                officeFirst = MeasureMillisecondsBatchValue(benchmark.OfficeIMO_WriteDataReader, invocationsPerLeg, out _);
                officeSecond = MeasureMillisecondsBatchValue(benchmark.OfficeIMO_WriteDataReader, invocationsPerLeg, out _);
                sylvanSecond = MeasureMillisecondsBatchValue(benchmark.Sylvan_WriteDataReader, invocationsPerLeg, out _);
            }

            officeSamples[index] = (officeFirst + officeSecond) / 2d;
            sylvanSamples[index] = (sylvanFirst + sylvanSecond) / 2d;
            pairedRatios[index] = officeSamples[index] / sylvanSamples[index];
        }

        double officeMedian = Median(officeSamples);
        double sylvanMedian = Median(sylvanSamples);
        Console.WriteLine(
            $"Paired CSV DataReader write ({shape}, {warmupIterations} warmups, {iterations} ABBA samples, {invocationsPerLeg} invocations per leg, affinity {affinity}, priority {priority}): " +
            $"OfficeIMO median {officeMedian:F3} ms, Sylvan median {sylvanMedian:F3} ms, " +
            $"ratio of medians {officeMedian / sylvanMedian:F4}, paired ratio median {Median(pairedRatios):F4} " +
            $"(P25 {Percentile(pairedRatios, 0.25d):F4}, P75 {Percentile(pairedRatios, 0.75d):F4}).");
    }
    return;
}

if (compareTypedParallelPaired) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 30;
    if (iterations <= 0) throw new ArgumentOutOfRangeException(nameof(iterations));
    string affinity = ApplyProcessAffinity(args, argumentIndex: 2);
    int? officeBatchSize = null;
    if (args.Length > 3 &&
        !string.Equals(args[3], "adaptive", StringComparison.OrdinalIgnoreCase)) {
        if (!int.TryParse(args[3], out int parsedBatchSize)) {
            throw new ArgumentException("OfficeIMO batch size must be a positive integer or 'adaptive'.");
        }
        officeBatchSize = parsedBatchSize;
    }
    int officeDegree = args.Length > 4 && int.TryParse(args[4], out int parsedOfficeDegree) ? parsedOfficeDegree : Environment.ProcessorCount;
    int sepDegree = args.Length > 5 && int.TryParse(args[5], out int parsedSepDegree) ? parsedSepDegree : Environment.ProcessorCount;
    int rowCount = args.Length > 6 && int.TryParse(args[6], out int parsedRowCount) ? parsedRowCount : 100_000;
    string priority = ApplyProcessPriority(args, argumentIndex: 7);
    if (officeBatchSize is <= 0) throw new ArgumentOutOfRangeException(nameof(officeBatchSize));
    if (officeDegree <= 0) throw new ArgumentOutOfRangeException(nameof(officeDegree));
    if (sepDegree <= 0) throw new ArgumentOutOfRangeException(nameof(sepDegree));
    if (rowCount <= 0) throw new ArgumentOutOfRangeException(nameof(rowCount));

    const int warmupIterations = 3;
    const int invocationsPerSample = 2;
    var benchmark = CsvTypedMaterializationFixture.Create(rowCount);
    Func<CsvBenchmarkRow[]> runOffice = () => benchmark.OfficeIMORecordParallel(officeDegree, officeBatchSize);
    Func<CsvBenchmarkRow[]> runSep = () => benchmark.SepParallel(sepDegree);
    benchmark.ValidateParallel(runOffice(), runSep());
    for (int index = 0; index < warmupIterations; index++) {
        runOffice();
        runSep();
    }

    var officeSamples = new double[iterations];
    var sepSamples = new double[iterations];
    var ratios = new double[iterations];
    var officeCpuSamples = new double[iterations];
    var sepCpuSamples = new double[iterations];
    var cpuRatios = new double[iterations];
    using var measuredProcess = System.Diagnostics.Process.GetCurrentProcess();
    for (int index = 0; index < iterations; index++) {
        CsvBenchmarkRow[] officeRows;
        CsvBenchmarkRow[] sepRows;
        if ((index & 1) == 0) {
            var officeFirst = MeasureValue(runOffice, measuredProcess, out officeRows);
            var sepFirst = MeasureValue(runSep, measuredProcess, out sepRows);
            var sepSecond = MeasureValue(runSep, measuredProcess, out sepRows);
            var officeSecond = MeasureValue(runOffice, measuredProcess, out officeRows);
            officeSamples[index] = (officeFirst.WallMilliseconds + officeSecond.WallMilliseconds) / invocationsPerSample;
            sepSamples[index] = (sepFirst.WallMilliseconds + sepSecond.WallMilliseconds) / invocationsPerSample;
            officeCpuSamples[index] = (officeFirst.CpuMilliseconds + officeSecond.CpuMilliseconds) / invocationsPerSample;
            sepCpuSamples[index] = (sepFirst.CpuMilliseconds + sepSecond.CpuMilliseconds) / invocationsPerSample;
        } else {
            var sepFirst = MeasureValue(runSep, measuredProcess, out sepRows);
            var officeFirst = MeasureValue(runOffice, measuredProcess, out officeRows);
            var officeSecond = MeasureValue(runOffice, measuredProcess, out officeRows);
            var sepSecond = MeasureValue(runSep, measuredProcess, out sepRows);
            officeSamples[index] = (officeFirst.WallMilliseconds + officeSecond.WallMilliseconds) / invocationsPerSample;
            sepSamples[index] = (sepFirst.WallMilliseconds + sepSecond.WallMilliseconds) / invocationsPerSample;
            officeCpuSamples[index] = (officeFirst.CpuMilliseconds + officeSecond.CpuMilliseconds) / invocationsPerSample;
            sepCpuSamples[index] = (sepFirst.CpuMilliseconds + sepSecond.CpuMilliseconds) / invocationsPerSample;
        }
        benchmark.ValidateParallel(officeRows, sepRows);
        ratios[index] = officeSamples[index] / sepSamples[index];
        cpuRatios[index] = officeCpuSamples[index] / sepCpuSamples[index];
    }

    Console.WriteLine(
        $"Paired typed parallel comparison ({rowCount} rows, {benchmark.TextLength} chars, {warmupIterations} warmups, {iterations} symmetric {invocationsPerSample}-invocation ABBA/BAAB samples, affinity {affinity}, priority {priority}): " +
        $"OfficeIMO batch {(officeBatchSize?.ToString() ?? "adaptive")}, DOP {officeDegree}: {Median(officeSamples):F3} ms; " +
        $"Sep DOP {sepDegree}: {Median(sepSamples):F3} ms; " +
        $"OfficeIMO/Sep wall paired median {Median(ratios):F4} (P25 {Percentile(ratios, 0.25d):F4}, P75 {Percentile(ratios, 0.75d):F4}); " +
        $"process CPU OfficeIMO {Median(officeCpuSamples):F3} ms, Sep {Median(sepCpuSamples):F3} ms, ratio {Median(cpuRatios):F4} " +
        $"(P25 {Percentile(cpuRatios, 0.25d):F4}, P75 {Percentile(cpuRatios, 0.75d):F4}).");
    return;
}

if (compareTypedSequentialPaired) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 20;
    if (iterations <= 0) {
        throw new ArgumentOutOfRangeException(nameof(iterations));
    }
    string affinity = ApplyProcessAffinity(args, argumentIndex: 2);
    const int warmupIterations = 3;
    const int invocationsPerSample = 4;

    var benchmark = CsvTypedMaterializationFixture.Create(100_000);
    for (int index = 0; index < warmupIterations; index++) {
        benchmark.OfficeIMORowsAs();
        benchmark.OfficeIMOManual();
        benchmark.SepSequential();
    }

    var rowsAsSamples = new double[iterations];
    var manualSamples = new double[iterations];
    var sepSamples = new double[iterations];
    var rowsAsRatios = new double[iterations];
    var manualRatios = new double[iterations];
    for (int index = 0; index < iterations; index++) {
        CsvBenchmarkRow[] rowsAsRows;
        CsvBenchmarkRow[] manualRows;
        CsvBenchmarkRow[] sepRows;
        switch (index % 3) {
            case 0:
                rowsAsSamples[index] = MeasureMillisecondsBatchValue(benchmark.OfficeIMORowsAs, invocationsPerSample, out rowsAsRows);
                manualSamples[index] = MeasureMillisecondsBatchValue(benchmark.OfficeIMOManual, invocationsPerSample, out manualRows);
                sepSamples[index] = MeasureMillisecondsBatchValue(benchmark.SepSequential, invocationsPerSample, out sepRows);
                break;
            case 1:
                manualSamples[index] = MeasureMillisecondsBatchValue(benchmark.OfficeIMOManual, invocationsPerSample, out manualRows);
                sepSamples[index] = MeasureMillisecondsBatchValue(benchmark.SepSequential, invocationsPerSample, out sepRows);
                rowsAsSamples[index] = MeasureMillisecondsBatchValue(benchmark.OfficeIMORowsAs, invocationsPerSample, out rowsAsRows);
                break;
            default:
                sepSamples[index] = MeasureMillisecondsBatchValue(benchmark.SepSequential, invocationsPerSample, out sepRows);
                rowsAsSamples[index] = MeasureMillisecondsBatchValue(benchmark.OfficeIMORowsAs, invocationsPerSample, out rowsAsRows);
                manualSamples[index] = MeasureMillisecondsBatchValue(benchmark.OfficeIMOManual, invocationsPerSample, out manualRows);
                break;
        }

        benchmark.ValidateSequential(rowsAsRows, manualRows, sepRows);
        rowsAsRatios[index] = rowsAsSamples[index] / sepSamples[index];
        manualRatios[index] = manualSamples[index] / sepSamples[index];
    }

    Console.WriteLine(
        $"Paired typed sequential comparison ({warmupIterations} warmups, {iterations} rotating {invocationsPerSample}-invocation samples, affinity {affinity}): " +
        $"OfficeIMO RowsAs {Median(rowsAsSamples):F3} ms, OfficeIMO manual {Median(manualSamples):F3} ms, Sep {Median(sepSamples):F3} ms; " +
        $"RowsAs/Sep paired median {Median(rowsAsRatios):F4} (P25 {Percentile(rowsAsRatios, 0.25d):F4}, P75 {Percentile(rowsAsRatios, 0.75d):F4}); " +
        $"manual/Sep paired median {Median(manualRatios):F4} (P25 {Percentile(manualRatios, 0.25d):F4}, P75 {Percentile(manualRatios, 0.75d):F4}).");
    return;
}

if (compareTrimUnescapeStringsPaired) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 100;
    if (iterations <= 0) {
        throw new ArgumentOutOfRangeException(nameof(iterations));
    }
    string affinity = ApplyProcessAffinity(args, argumentIndex: 2);
    const int warmupIterations = 3;
    const int invocationsPerSample = 32;

    var benchmark = new CsvTrimUnescapeBenchmarks { RowCount = 50_000 };
    benchmark.Setup();
    for (int index = 0; index < warmupIterations; index++) {
        MeasureMillisecondsBatch(benchmark.OfficeIMODataReaderStrings, invocationsPerSample, out _);
        MeasureMillisecondsBatch(benchmark.SepStrings, invocationsPerSample, out _);
    }

    var officeSamples = new double[iterations];
    var sepSamples = new double[iterations];
    var ratios = new double[iterations];
    for (int index = 0; index < iterations; index++) {
        CsvReadObservation officeObservation;
        CsvReadObservation sepObservation;
        if ((index & 1) == 0) {
            officeSamples[index] = MeasureMillisecondsBatch(benchmark.OfficeIMODataReaderStrings, invocationsPerSample, out officeObservation);
            sepSamples[index] = MeasureMillisecondsBatch(benchmark.SepStrings, invocationsPerSample, out sepObservation);
        } else {
            sepSamples[index] = MeasureMillisecondsBatch(benchmark.SepStrings, invocationsPerSample, out sepObservation);
            officeSamples[index] = MeasureMillisecondsBatch(benchmark.OfficeIMODataReaderStrings, invocationsPerSample, out officeObservation);
        }

        if (officeObservation != sepObservation) {
            throw new InvalidDataException(
                $"Paired trim/unescape sample {index} produced different observations: OfficeIMO={officeObservation}; Sep={sepObservation}.");
        }

        ratios[index] = officeSamples[index] / sepSamples[index];
    }

    Console.WriteLine(
        $"Paired trim/unescape string comparison ({warmupIterations} warmup blocks, {iterations} rotating {invocationsPerSample}-invocation samples, affinity {affinity}, " +
        $"AVX512BW={Avx512BW.IsSupported}, AVX2={Avx2.IsSupported}): " +
        $"OfficeIMO {Median(officeSamples):F3} ms, Sep {Median(sepSamples):F3} ms; " +
        $"OfficeIMO/Sep paired median {Median(ratios):F4} " +
        $"(P25 {Percentile(ratios, 0.25d):F4}, P75 {Percentile(ratios, 0.75d):F4}).");
    return;
}

if (comparePaired) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 100;
    if (iterations <= 0) {
        throw new ArgumentOutOfRangeException(nameof(iterations));
    }
    string affinity = ApplyProcessAffinity(args, argumentIndex: 2);
    const int warmupIterations = 10;

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

if (profileTypedOfficeRowsAs || profileTypedOfficeManual || profileTypedOfficeParallel || profileTypedSepSequential || profileTypedSepParallel) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 50;
    if (iterations <= 0) {
        throw new ArgumentOutOfRangeException(nameof(iterations));
    }
    ApplyProcessAffinity(args, argumentIndex: 2);
    int parallelBatchSize = args.Length > 3 && int.TryParse(args[3], out int parsedBatchSize)
        ? parsedBatchSize
        : 2048;
    int parallelDegree = args.Length > 4 && int.TryParse(args[4], out int parsedDegree)
        ? parsedDegree
        : Environment.ProcessorCount;

    var benchmark = CsvTypedMaterializationFixture.Create(100_000);
    Func<CsvBenchmarkRow[]> run;
    string implementation;
    if (profileTypedOfficeRowsAs) {
        run = benchmark.OfficeIMORowsAs;
        implementation = "OfficeIMO RowsAs";
    } else if (profileTypedOfficeManual) {
        run = benchmark.OfficeIMOManual;
        implementation = "OfficeIMO manual";
    } else if (profileTypedOfficeParallel) {
        run = () => benchmark.OfficeIMORecordParallel(parallelDegree, parallelBatchSize);
        implementation = "OfficeIMO transient-record parallel";
    } else if (profileTypedSepSequential) {
        run = benchmark.SepSequential;
        implementation = "Sep sequential";
    } else {
        run = () => benchmark.SepParallel(parallelDegree);
        implementation = "Sep parallel";
    }

    for (int index = 0; index < 3; index++) {
        run();
    }

    CsvBenchmarkRow[] rows = [];
    var stopwatch = System.Diagnostics.Stopwatch.StartNew();
    for (int index = 0; index < iterations; index++) {
        rows = run();
    }
    stopwatch.Stop();

    Console.WriteLine(
        $"Profiled {implementation} typed materialization {iterations} times in {stopwatch.Elapsed.TotalMilliseconds:F2} ms " +
        $"({stopwatch.Elapsed.TotalMilliseconds / iterations:F3} ms/iteration): {rows.Length} rows.");
    return;
}

if (profileTrimSpanOfficeIMO || profileTrimSpanSep) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 100;
    if (iterations <= 0) {
        throw new ArgumentOutOfRangeException(nameof(iterations));
    }
    ApplyProcessAffinity(args, argumentIndex: 2);

    var benchmark = new CsvTrimUnescapeSpanBenchmarks { RowCount = 50_000 };
    benchmark.Setup();
    Func<CsvReadObservation> run = profileTrimSpanOfficeIMO
        ? benchmark.OfficeIMOFieldSpans
        : benchmark.SepSpans;
    string implementation = profileTrimSpanOfficeIMO ? "OfficeIMO" : "Sep";
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
        $"Profiled {implementation} trim/unescape spans {iterations} times in {stopwatch.Elapsed.TotalMilliseconds:F2} ms " +
        $"({stopwatch.Elapsed.TotalMilliseconds / iterations:F3} ms/iteration): {observation}.");
    return;
}

if (profileTrimOfficeIMO || profileTrimSep) {
    int iterations = args.Length > 1 && int.TryParse(args[1], out int parsedIterations)
        ? parsedIterations
        : 100;
    if (iterations <= 0) {
        throw new ArgumentOutOfRangeException(nameof(iterations));
    }
    ApplyProcessAffinity(args, argumentIndex: 2);

    var benchmark = new CsvTrimUnescapeBenchmarks { RowCount = 50_000 };
    benchmark.Setup();
    Func<CsvReadObservation> run = profileTrimOfficeIMO
        ? benchmark.OfficeIMODataReaderStrings
        : benchmark.SepStrings;
    string implementation = profileTrimOfficeIMO ? "OfficeIMO" : "Sep";
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
        $"Profiled {implementation} trim/unescape strings {iterations} times in {stopwatch.Elapsed.TotalMilliseconds:F2} ms " +
        $"({stopwatch.Elapsed.TotalMilliseconds / iterations:F3} ms/iteration): {observation}.");
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

var (priorityArgs, benchmarkPriority) = ExtractBenchmarkPriority(args);
if (benchmarkPriority != null) {
    BenchmarkProcessorAffinity.ApplyPriority(benchmarkPriority);
}
var (benchmarkArgs, affinityMasks) = ExtractAffinityMasks(priorityArgs);
var config = ManualConfig
    .Create(DefaultConfig.Instance)
    .AddDiagnoser(MemoryDiagnoser.Default)
    .AddExporter(JsonExporter.Full)
    .WithSummaryStyle(SummaryStyle.Default.WithRatioStyle(RatioStyle.Percentage))
    .AddColumn(StatisticColumn.OperationsPerSecond);

for (int index = 0; index < affinityMasks.Length; index++) {
    IntPtr affinity = affinityMasks[index];
    Job job = Job.Default
        .WithAffinity(affinity)
        .WithId($"Affinity-{BenchmarkProcessorAffinity.Format(affinity)}");
    if (benchmarkPriority != null) {
        job = job.WithEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY", benchmarkPriority);
    }
    config.AddJob(job);
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(benchmarkArgs, config);

static double MeasureMilliseconds(
    Func<CsvReadObservation> operation,
    out CsvReadObservation observation) {
    long started = System.Diagnostics.Stopwatch.GetTimestamp();
    observation = operation();
    return System.Diagnostics.Stopwatch.GetElapsedTime(started).TotalMilliseconds;
}

static string ApplyProcessAffinity(string[] arguments, int argumentIndex)
    => arguments.Length <= argumentIndex
        ? "unchanged"
        : BenchmarkProcessorAffinity.Apply(arguments[argumentIndex]);

static string ApplyProcessPriority(string[] arguments, int argumentIndex) {
    if (arguments.Length <= argumentIndex ||
        string.Equals(arguments[argumentIndex], "unchanged", StringComparison.OrdinalIgnoreCase)) {
        return System.Diagnostics.Process.GetCurrentProcess().PriorityClass.ToString();
    }

    if (!Enum.TryParse(arguments[argumentIndex], ignoreCase: true, out System.Diagnostics.ProcessPriorityClass priority) ||
        priority == System.Diagnostics.ProcessPriorityClass.RealTime) {
        throw new ArgumentException("Priority must be Idle, BelowNormal, Normal, AboveNormal, or High.");
    }

    using System.Diagnostics.Process process = System.Diagnostics.Process.GetCurrentProcess();
    process.PriorityClass = priority;
    return process.PriorityClass.ToString();
}

static double MeasureMillisecondsBatch(
    Func<CsvReadObservation> operation,
    int invocationCount,
    out CsvReadObservation observation) {
    long started = System.Diagnostics.Stopwatch.GetTimestamp();
    observation = default;
    for (int index = 0; index < invocationCount; index++) {
        observation = operation();
    }
    return System.Diagnostics.Stopwatch.GetElapsedTime(started).TotalMilliseconds / invocationCount;
}

static double MeasureMillisecondsBatchValue<T>(
    Func<T> operation,
    int invocationCount,
    out T result) {
    long started = System.Diagnostics.Stopwatch.GetTimestamp();
    result = default!;
    for (int index = 0; index < invocationCount; index++) {
        result = operation();
    }
    return System.Diagnostics.Stopwatch.GetElapsedTime(started).TotalMilliseconds / invocationCount;
}

static (double WallMilliseconds, double CpuMilliseconds) MeasureBatchValue<T>(
    Func<T> operation,
    int invocationCount,
    System.Diagnostics.Process process,
    out T result) {
    TimeSpan cpuStarted = process.TotalProcessorTime;
    long started = System.Diagnostics.Stopwatch.GetTimestamp();
    result = default!;
    for (int index = 0; index < invocationCount; index++) {
        result = operation();
    }
    double wallMilliseconds = System.Diagnostics.Stopwatch.GetElapsedTime(started).TotalMilliseconds / invocationCount;
    double cpuMilliseconds = (process.TotalProcessorTime - cpuStarted).TotalMilliseconds / invocationCount;
    return (wallMilliseconds, cpuMilliseconds);
}

static (double WallMilliseconds, double CpuMilliseconds) MeasureValue<T>(
    Func<T> operation,
    System.Diagnostics.Process process,
    out T result) {
    TimeSpan cpuStarted = process.TotalProcessorTime;
    long started = System.Diagnostics.Stopwatch.GetTimestamp();
    result = operation();
    double wallMilliseconds = System.Diagnostics.Stopwatch.GetElapsedTime(started).TotalMilliseconds;
    double cpuMilliseconds = (process.TotalProcessorTime - cpuStarted).TotalMilliseconds;
    return (wallMilliseconds, cpuMilliseconds);
}

static (string[] Arguments, IntPtr[] Masks) ExtractAffinityMasks(string[] arguments) {
    int optionIndex = Array.FindIndex(
        arguments,
        static argument => string.Equals(argument, "--affinityMasks", StringComparison.OrdinalIgnoreCase));
    if (optionIndex < 0) {
        return (arguments, Array.Empty<IntPtr>());
    }
    if (optionIndex + 1 >= arguments.Length) {
        throw new ArgumentException("--affinityMasks requires a comma-separated list of positive processor-affinity masks.");
    }

    IntPtr[] masks = BenchmarkProcessorAffinity.ParseList(arguments[optionIndex + 1]);

    var forwarded = new string[arguments.Length - 2];
    if (optionIndex > 0) {
        Array.Copy(arguments, 0, forwarded, 0, optionIndex);
    }
    if (optionIndex + 2 < arguments.Length) {
        Array.Copy(
            arguments,
            optionIndex + 2,
            forwarded,
            optionIndex,
            arguments.Length - optionIndex - 2);
    }
    return (forwarded, masks);
}

static (string[] Arguments, string? Priority) ExtractBenchmarkPriority(string[] arguments) {
    int optionIndex = Array.FindIndex(
        arguments,
        static argument => string.Equals(argument, "--priority", StringComparison.OrdinalIgnoreCase));
    if (optionIndex < 0) {
        return (arguments, null);
    }
    if (optionIndex + 1 >= arguments.Length) {
        throw new ArgumentException("--priority requires Idle, BelowNormal, Normal, AboveNormal, or High.");
    }

    string priority = arguments[optionIndex + 1];
    var forwarded = new string[arguments.Length - 2];
    if (optionIndex > 0) {
        Array.Copy(arguments, 0, forwarded, 0, optionIndex);
    }
    if (optionIndex + 2 < arguments.Length) {
        Array.Copy(arguments, optionIndex + 2, forwarded, optionIndex, arguments.Length - optionIndex - 2);
    }
    return (forwarded, priority);
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
