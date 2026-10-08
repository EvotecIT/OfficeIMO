using System.Reflection;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Exporters.Json;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;
using BenchmarkDotNet.Validators;
using OfficeIMO.Benchmarks;
using OfficeIMO.Excel.ReaderComparison.Benchmarks;
using OfficeIMO.Excel;

AssemblyMetadataAttribute[] metadata = Assembly.GetExecutingAssembly().GetCustomAttributes<AssemblyMetadataAttribute>().ToArray();
string? packageVersion = metadata.Single(attribute => attribute.Key == "OfficeIMOBenchmarkPackageVersion").Value;
string? assemblyDirectory = metadata.Single(attribute => attribute.Key == "OfficeIMOBenchmarkAssemblyDirectory").Value;
string? newApis = metadata.Single(attribute => attribute.Key == "OfficeIMOBenchmarkNewApis").Value;
string? csv = metadata.Single(attribute => attribute.Key == "OfficeIMOBenchmarkCsv").Value;
string? arrow = metadata.Single(attribute => attribute.Key == "OfficeIMOBenchmarkArrow").Value;
BenchmarkInput.WriteDescription();

if (args is ["--validate-bdn"]) {
    int typeCount = 0, methodCount = 0, caseCount = 0, errorCount = 0;
    foreach (Type type in Assembly.GetExecutingAssembly().GetTypes().Where(type =>
        type.GetMethods(BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic)
            .Any(method => method.IsDefined(typeof(BenchmarkAttribute), inherit: true)))) {
        BenchmarkRunInfo runInfo = BenchmarkConverter.TypeToBenchmarks(type);
        typeCount++;
        methodCount += runInfo.BenchmarksCases.Select(benchmark => benchmark.Descriptor.WorkloadMethod).Distinct().Count();
        caseCount += runInfo.BenchmarksCases.Length;
        foreach (ValidationError error in CompilationValidator.FailOnError.Validate(runInfo)) {
            Console.Error.WriteLine(error.Message);
            if (error.IsCritical) errorCount++;
        }
    }
    if (errorCount != 0) throw new InvalidOperationException($"BenchmarkDotNet rejected {errorCount} workload declarations.");
    Console.WriteLine($"BenchmarkDotNet declarations validated: {typeCount} types, {methodCount} methods, {caseCount} parameter cases.");
    return;
}

if (args is ["--validate-ado"]) {
    foreach (int rows in new AdoReadBenchmarks().RowCounts())
        foreach (AdoAccess access in Enum.GetValues<AdoAccess>())
            await new AdoReadBenchmarks { RowCount = rows, Access = access }.SetupAsync();
    foreach (int rows in new AdoReadBenchmarks().RowCounts())
        await new DataTableLoadDiagnostics { RowCount = rows }.SetupAsync();
    Console.WriteLine("Validated every ADO access mode.");
    return;
}

if (args is ["--validate-cold"]) {
    var reader = new ColdStartReadBenchmarks();
    await reader.SetupAsync();
    reader.Validate();
    var writer = new ColdStartWriteBenchmarks();
    writer.Setup();
    await writer.ValidateAsync();
    Console.WriteLine("Validated cold-use contracts in this qualification process; measurements use fresh processes.");
    return;
}

if (args is ["--validate-models"]) {
    foreach (int rows in new TypedModelReadBenchmarks().RowCounts()) {
        foreach (TypedModelKind model in Enum.GetValues<TypedModelKind>())
            await new TypedModelReadBenchmarks { RowCount = rows, Model = model }.SetupAsync();
#if OFFICEIMO_BENCHMARK_NEW_APIS
        await new BorrowedRefStructReadBenchmarks { RowCount = rows }.SetupAsync();
#endif
    }
    return;
}

if (args is ["--validate-real-typed"]) {
    foreach (ComparisonWorkbookFormat format in new[] { ComparisonWorkbookFormat.Xlsx, ComparisonWorkbookFormat.Xlsb })
        new RealDataTypedReadBenchmarks { Format = format }.Setup();
    return;
}

if (args is ["--validate-shared-stages"]) {
    foreach (int rows in new StringHeavyReadBenchmarks().RowCounts())
        foreach (SharedStringZipStorage storage in Enum.GetValues<SharedStringZipStorage>())
            await new SharedStringFirstRowBenchmarks { RowCount = rows, Storage = storage }.SetupAsync();
    return;
}

#if OFFICEIMO_BENCHMARK_CSV
#if OFFICEIMO_BENCHMARK_NEW_APIS
if (args is ["--validate-csv-direct-aggregate-scale"]) {
    await CsvBenchmarkQualification.ValidateDirectAggregateScaleAsync();
    return;
}
if (args is ["--validate-csv-direct-aggregate"]) {
    await CsvBenchmarkQualification.ValidateDirectAggregateAsync();
    return;
}
#endif
if (args is ["--validate-real-csv"]) {
    new CsvRealDataReadBenchmarks().Setup();
    return;
}
if (args is ["--validate-csv-record-write"]) {
    foreach (int rows in new TypedReadBenchmarks().RowCounts())
        foreach (bool mapped in new[] { false, true })
            await new CsvRecordWriterBenchmarks { RowCount = rows, Mapped = mapped }.SetupAsync();
    return;
}
#if OFFICEIMO_BENCHMARK_NEW_APIS
if (args is ["--validate-csv-utf8-writer"]) {
    CsvBenchmarkQualification.ValidateUtf8Writer();
    return;
}
#endif
if (args is ["--validate-csv"]) {
    await CsvBenchmarkQualification.ValidateAsync();
    return;
}
if (args is ["--validate-csv-parallel"]) {
    await CsvBenchmarkQualification.ValidateParallelAsync();
    return;
}
#endif

#if OFFICEIMO_BENCHMARK_ARROW && OFFICEIMO_BENCHMARK_NEW_APIS
if (args is ["--validate-arrow"]) {
    await ArrowBenchmarkValidation.RunAsync();
    return;
}
if (args is ["--validate-write-arrow-strings"]) {
    var workload = new ArrowStringWriterBenchmarks();
    try { workload.Setup(); } finally { workload.Cleanup(); }
    return;
}
#endif

if (args is ["--validate"]) {
    var matrix = new TypedReadBenchmarks();
    foreach (int rows in matrix.RowCounts()) {
        foreach (string shape in matrix.Shapes()) {
            var workload = new TypedReadBenchmarks { RowCount = rows, Shape = shape };
            await workload.SetupAsync();
        }
        await new TypedXlsbReadBenchmarks { RowCount = rows }.SetupAsync();
    }
    return;
}

if (args is ["--validate-raw"]) {
    var matrix = new RawReadBenchmarks();
    foreach (int rows in matrix.RowCounts()) {
        foreach (RawWorkbookFormat format in Enum.GetValues<RawWorkbookFormat>())
            await new RawReadBenchmarks { RowCount = rows, Format = format }.SetupAsync();
    }
    return;
}

if (args is ["--validate-async"]) {
    var matrix = new TypedAsyncReadBenchmarks();
    foreach (int rows in matrix.RowCounts()) {
        foreach (string shape in matrix.Shapes())
            await new TypedAsyncReadBenchmarks { RowCount = rows, Shape = shape }.SetupAsync();
        foreach (RawWorkbookFormat format in Enum.GetValues<RawWorkbookFormat>())
            await new RawAsyncReadBenchmarks { RowCount = rows, Format = format }.SetupAsync();
        await new TypedXlsbAsyncReadBenchmarks { RowCount = rows }.SetupAsync();
    }
    return;
}

if (args is ["--validate-write"]) {
    var matrix = new WriterBenchmarks();
    foreach (int rows in matrix.RowCounts()) await new WriterBenchmarks { RowCount = rows }.SetupAsync();
    return;
}

if (args is ["--validate-write-options"]) {
    foreach (int rows in new WriterBenchmarks().RowCounts()) {
        await new SharedStringWriterBenchmarks { RowCount = rows }.SetupAsync();
        await new CompactWriterBenchmarks { RowCount = rows }.SetupAsync();
        await new NativeXlsWriterDiagnosticBenchmarks { RowCount = rows }.SetupAsync();
        await new ConfiguredXlsbWriterBenchmarks { RowCount = rows }.SetupAsync();
    }
    return;
}

if (args is ["--validate-write-xlsb-options"]) {
    foreach (int rows in new WriterBenchmarks().RowCounts())
        await new ConfiguredXlsbWriterBenchmarks { RowCount = rows }.SetupAsync();
    return;
}

if (args is ["--validate-write-records"]) {
    foreach (int rows in new WriterBenchmarks().RowCounts()) {
        foreach (ExcelFileFormat format in new[] { ExcelFileFormat.Xlsx, ExcelFileFormat.Xlsb, ExcelFileFormat.Xls })
            await new RecordWriterBenchmarks { RowCount = rows, Format = format }.SetupAsync();
    }
    return;
}

#if OFFICEIMO_BENCHMARK_NEW_APIS
if (args is ["--validate-write-styled-rows"]) {
    var workload = new StyledRowWriterBenchmarks();
    try { workload.Setup(); workload.SaveQualifiedArtifacts(); } finally { workload.Cleanup(); }
    return;
}
if (args is ["--validate-shared-utf8"]) {
    foreach (int rows in new SharedStringUtf8ReadBenchmarks().RowCounts())
        foreach (ComparisonWorkbookFormat format in new[] { ComparisonWorkbookFormat.Xlsx, ComparisonWorkbookFormat.Xlsb })
            foreach (SharedStringZipStorage storage in Enum.GetValues<SharedStringZipStorage>())
                await new SharedStringUtf8ReadBenchmarks { RowCount = rows, Format = format, Storage = storage }.SetupAsync();
    return;
}
if (args is ["--prepare-encrypted"]) {
    await EncryptedWorkbookFixture.PrepareLargeAsync();
    return;
}
if (args is ["--validate-encrypted"]) {
    foreach (EncryptedWorkbookInput input in Enum.GetValues<EncryptedWorkbookInput>())
        await new EncryptedVerifiedReadBenchmarks { Input = input }.SetupAsync();
    await new EncryptedOriginalSmallDiagnostics().SetupAsync();
    await new EncryptedOriginalLargeDiagnostics().SetupAsync();
    return;
}
if (args is ["--validate-stream-async"]) {
    foreach (int rows in new TypedStreamAsyncReadBenchmarks().RowCounts())
        foreach (string shape in new TypedStreamAsyncReadBenchmarks().Shapes())
            await new TypedStreamAsyncReadBenchmarks { RowCount = rows, Shape = shape }.SetupAsync();
    return;
}
if (args is ["--validate-write-utf8"]) {
    var workload = new Utf8WriterBenchmarks();
    try { workload.Setup(); } finally { workload.Cleanup(); }
    return;
}
#endif

if (args is ["--validate-real"]) {
    foreach (ComparisonWorkbookFormat format in Enum.GetValues<ComparisonWorkbookFormat>())
        new RealDataReadBenchmarks { Format = format }.Setup();
    return;
}

if (args is ["--validate-strings"]) {
    foreach (int rows in new StringHeavyReadBenchmarks().RowCounts()) {
        foreach (ComparisonWorkbookFormat format in new[] { ComparisonWorkbookFormat.Xlsx, ComparisonWorkbookFormat.Xlsb })
            await new StringHeavyReadBenchmarks { RowCount = rows, Format = format }.SetupAsync();
    }
    return;
}

if (args.Length is 1 or 2 && args[0] == "--profile-officeimo-write") {
    int iterations = args.Length == 1 ? 500 : int.Parse(args[1], System.Globalization.CultureInfo.InvariantCulture);
    if (iterations is < 1 or > 100_000) throw new ArgumentOutOfRangeException(nameof(iterations));
    var workload = new WriterBenchmarks();
    await workload.SetupAsync();
    for (int warmup = 0; warmup < 50; warmup++) workload.OfficeIMO();
    Console.WriteLine($"CPU profile workload: {iterations} complete OfficeIMO writes after 50 warmups.");
    long observation = 0;
    for (int iteration = 0; iteration < iterations; iteration++) observation = unchecked(observation + workload.OfficeIMO());
    Console.WriteLine($"CPU profile workload completed; aggregate package length={observation}.");
    return;
}

if (args.Length is 1 or 2 && args[0] == "--profile-officeimo") {
    int iterations = args.Length == 1 ? 500 : int.Parse(args[1], System.Globalization.CultureInfo.InvariantCulture);
    if (iterations is < 1 or > 100_000) throw new ArgumentOutOfRangeException(nameof(iterations));
    var workload = new TypedReadBenchmarks();
    await workload.SetupAsync();
    for (int warmup = 0; warmup < 50; warmup++) workload.OfficeIMOTyped();
    Console.WriteLine($"CPU profile workload: {iterations} complete OfficeIMO typed scans after 50 warmups.");
    long observation = 0;
    for (int iteration = 0; iteration < iterations; iteration++)
        observation = unchecked(observation + workload.OfficeIMOTyped());
    Console.WriteLine($"CPU profile workload completed; aggregate observation={observation}.");
    return;
}

// Preserve the explicit package choice in BDN's generated build; otherwise its
// ProjectReference would silently compile the default source lane instead.
var buildArguments = new List<Argument> { new MsBuildArgument("/p:UseSharedCompilation=false") };
if (!string.IsNullOrEmpty(packageVersion))
    buildArguments.Add(new MsBuildArgument($"/p:OfficeIMOBenchmarkPackageVersion={packageVersion}"));
if (!string.IsNullOrEmpty(assemblyDirectory)) {
    buildArguments.Add(new MsBuildArgument($"/p:OfficeIMOBenchmarkAssemblyDirectory=\"{assemblyDirectory}\""));
    string savedReferenceTargets = Path.Combine(AppContext.BaseDirectory, "SavedAssemblyReferences.targets");
    if (!File.Exists(savedReferenceTargets))
        throw new FileNotFoundException("The saved-assembly benchmark lane requires SavedAssemblyReferences.targets.", savedReferenceTargets);
    buildArguments.Add(new MsBuildArgument($"/p:CustomAfterMicrosoftCommonTargets=\"{savedReferenceTargets}\""));
}
if (string.Equals(newApis, "true", StringComparison.OrdinalIgnoreCase))
    buildArguments.Add(new MsBuildArgument("/p:OfficeIMOBenchmarkNewApis=true"));
if (string.Equals(csv, "true", StringComparison.OrdinalIgnoreCase))
    buildArguments.Add(new MsBuildArgument("/p:OfficeIMOBenchmarkCsv=true"));
if (string.Equals(arrow, "true", StringComparison.OrdinalIgnoreCase))
    buildArguments.Add(new MsBuildArgument("/p:OfficeIMOBenchmarkArrow=true"));
var benchmarkJob = Job.Default.WithArguments(buildArguments.ToArray());
string? affinityMask = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_AFFINITY_MASK");
if (!string.IsNullOrEmpty(affinityMask)) {
    if (args.Any(argument => argument == "--affinity" || argument.StartsWith("--affinity=", StringComparison.Ordinal)))
        throw new ArgumentException("Choose OFFICEIMO_BENCHMARK_AFFINITY_MASK or --affinity, rather than both.");
    IntPtr[] masks = BenchmarkProcessorAffinity.ParseList(affinityMask);
    if (masks.Length != 1)
        throw new ArgumentException("OFFICEIMO_BENCHMARK_AFFINITY_MASK requires exactly one processor mask.");
    benchmarkJob = benchmarkJob.WithAffinity(masks[0]);
}
var config = ManualConfig.Create(DefaultConfig.Instance)
    .AddExporter(JsonExporter.Full)
    .AddJob(benchmarkJob.AsMutator());
BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args, config);
