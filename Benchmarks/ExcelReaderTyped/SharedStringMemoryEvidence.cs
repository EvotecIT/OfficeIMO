using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Reflection;
using System.Runtime;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Data;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Opt-in retained-memory observations around one completely qualified XLSX scan.</summary>
internal static class SharedStringMemoryEvidence {
    private enum TextPolicy { Materialized, Utf8 }

    internal static async Task SaveFixtureAsync(string path) {
        int rows = GetSingleRowCount();
        byte[] source = await StringHeavyWorkbookGenerator.BuildXlsxAsync(rows);
        using (var input = new MemoryStream(source, writable: false)) {
            using var package = new System.IO.Compression.ZipArchive(input, System.IO.Compression.ZipArchiveMode.Read);
            WrittenWorkbookValidation.ValidateEntries(package);
            WrittenWorkbookValidation.ValidateRelationships(package);
        }
        using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write)) output.Write(source);
        BenchmarkInput.WriteWorkbookFixtureIdentity($"shared-memory-frozen/Xlsx/dataRows={rows}", source);
    }

    internal static async Task RunAsync(string[] args) {
        if (args.Length != 2)
            throw new ArgumentException("Usage: --measure-shared-memory <Stored|Deflated> <Materialized|Utf8>. Run each policy in a fresh process.");
        SharedStringZipStorage storage = args[0].ToUpperInvariant() switch {
            "STORED" => SharedStringZipStorage.Stored,
            "DEFLATED" => SharedStringZipStorage.Deflated,
            _ => throw new ArgumentException("Shared-string ZIP storage must be Stored or Deflated."),
        };
        TextPolicy policy = args[1].ToUpperInvariant() switch {
            "MATERIALIZED" => TextPolicy.Materialized,
            "UTF8" => TextPolicy.Utf8,
            _ => throw new ArgumentException("Shared-string text policy must be Materialized or Utf8."),
        };
#if !OFFICEIMO_BENCHMARK_NEW_APIS
        if (policy == TextPolicy.Utf8)
            throw new InvalidOperationException("Utf8 memory evidence requires -p:OfficeIMOBenchmarkNewApis=true.");
#endif
        int rows = GetSingleRowCount();

        // Do not call a reader qualification/setup method here: it would populate
        // reader caches and pools before the baseline observation.
        Fixture fixture = await BuildFixtureAsync(rows, storage);
        BenchmarkInput.WriteWorkbookFixtureIdentity($"shared-memory/Xlsx/{storage}/dataRows={rows}", fixture.Bytes);
        Assembly assembly = typeof(ExcelDocument).Assembly;
        string assemblySha256;
        using (Stream input = File.OpenRead(assembly.Location))
            assemblySha256 = Convert.ToHexString(SHA256.HashData(input));
        MemorySnapshot[] snapshots = Measure(fixture.Bytes, rows, policy);

        var report = new {
            kind = "SharedStringMemoryEvidence",
            format = "Xlsx",
            storage = storage.ToString(),
            textPolicy = policy.ToString(),
            dataRows = rows,
            rowsIncludingHeader = rows + 1,
            fieldsPerRow = StringHeavyWorkbookGenerator.Headers.Length,
            validatedFields = (long)(rows + 1) * StringHeavyWorkbookGenerator.Headers.Length,
            validation = "Every header/text byte or string, native Double/DateTime value and typed getter, row order/count, raw object schema and single sheet.",
            fixtureBytes = fixture.Bytes.LongLength,
            sourceFixturePolicy = fixture.FromSavedInput ? "SavedInputLoadedBeforeBaseline" : "GeneratedBeforeBaseline",
            sourceFixtureSha256 = fixture.SourceSha256,
            fixtureSha256 = fixture.Sha256,
            officeIMOExcelInformationalVersion = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion,
            officeIMOExcelAssemblySha256 = assemblySha256,
            processId = Environment.ProcessId,
            runtime = RuntimeInformation.FrameworkDescription,
            operatingSystem = RuntimeInformation.OSDescription,
            serverGc = GCSettings.IsServerGC,
            managedScope = "GC-retained managed bytes after forced full collections; deltas from before opening include existing pool/static-cache retention and validation effects, not allocation totals or an isolated reader footprint.",
            processScope = "Current working set and supported .NET peak working set belong to this complete process, including fixture generation, package qualification and JIT; peaks are not per-reader peaks. A null peak means unavailable; on macOS wrap the fresh process with /usr/bin/time -l for lifetime maximum RSS.",
            snapshots,
        };
        Console.WriteLine($"Qualified shared-string memory evidence: {storage}/{policy}, {rows + 1} rows, "
            + $"{report.validatedFields} fields; source SHA256={fixture.SourceSha256}; fixture SHA256={fixture.Sha256}.");
        Console.WriteLine(JsonSerializer.Serialize(report, new JsonSerializerOptions {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        }));
    }

    private static async Task<Fixture> BuildFixtureAsync(int rows, SharedStringZipStorage storage) {
        string? configured = Environment.GetEnvironmentVariable("OFFICEIMO_SHARED_XLSX_FIXTURE");
        byte[] source = configured == null ? await StringHeavyWorkbookGenerator.BuildXlsxAsync(rows) : File.ReadAllBytes(configured);
        string sourceSha256 = Convert.ToHexString(SHA256.HashData(source));
        byte[] bytes = storage == SharedStringZipStorage.Stored ? SharedStringLifecycleWorkload.Store(source) : source;
        if (storage == SharedStringZipStorage.Stored)
            SharedStringLifecycleWorkload.ValidateInflatedParts(source, bytes);
        using (var stream = new MemoryStream(bytes, writable: false)) {
            using var package = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Read);
            if (storage == SharedStringZipStorage.Deflated) WrittenWorkbookValidation.ValidateEntries(package);
            WrittenWorkbookValidation.ValidateRelationships(package);
        }
        return new Fixture(bytes, sourceSha256, Convert.ToHexString(SHA256.HashData(bytes)), configured != null);
    }

    private static int GetSingleRowCount() {
        int[] counts = new StringHeavyReadBenchmarks().RowCounts().ToArray();
        return counts.Length == 1 ? counts[0]
            : throw new ArgumentException("Memory evidence requires exactly one OFFICEIMO_STRING_BENCHMARK_ROWS count per fresh process.");
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static MemorySnapshot[] Measure(byte[] bytes, int rows, TextPolicy policy) {
        using Process process = Process.GetCurrentProcess();
        MemorySnapshot beforeOpen = Snapshot(process, "BeforeOpen");
        // Returning from this separate frame releases the disposed reader and all
        // temporary expected values before the final full-GC observation.
        (MemorySnapshot afterHeader, MemorySnapshot afterScan) = ScanAndDispose(bytes, rows, policy, process, beforeOpen.ManagedRetainedBytes);
        MemorySnapshot afterDisposal = Snapshot(process, "AfterDisposal", beforeOpen.ManagedRetainedBytes);
        GC.KeepAlive(bytes); // The same fixture buffer is retained at every stage.
        return [beforeOpen, afterHeader, afterScan, afterDisposal];
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static (MemorySnapshot Header, MemorySnapshot Scan) ScanAndDispose(
        byte[] bytes, int rows, TextPolicy policy, Process process, long baseline) {
        using var reader = ExcelDocument.OpenDataReader(bytes, new ExcelReadOptions { HasHeaderRow = false });
        if (reader.SheetNames.Count != 1 || reader.CurrentSheetName != "S1"
            || reader.CurrentSheetIndex != 0 || reader.CurrentResultIndex != 0 || !reader.Read())
            throw new InvalidDataException("The SST memory fixture must have one S1 sheet and its header row.");
        ValidateRow(reader, StringHeavyWorkbookGenerator.Headers.Cast<object>().ToArray(), policy, 0);
        ValidateSchema(reader);
        MemorySnapshot afterHeader = Snapshot(process, "AfterFirstHeader", baseline);
        GC.KeepAlive(reader);

        ValidateDataRows(reader, rows, policy);
        ValidateSchema(reader);
        MemorySnapshot afterScan = Snapshot(process, "AfterFullScanReaderAlive", baseline);
        GC.KeepAlive(reader);
        if (reader.NextResult()) throw new InvalidDataException("The SST memory fixture has an extra sheet.");
        return (afterHeader, afterScan);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void ValidateDataRows(ExcelWorkbookDataReader reader, int rows, TextPolicy policy) {
        int count = 0;
        while (reader.Read()) {
            if (++count > rows) throw new InvalidDataException("The SST memory fixture has extra rows.");
            ValidateRow(reader, StringHeavyWorkbookGenerator.ExpectedValues(count), policy, count);
        }
        if (count != rows) throw new InvalidDataException($"The SST memory scan returned {count} data rows instead of {rows}.");
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void ValidateRow(ExcelWorkbookDataReader reader, object[] expected, TextPolicy policy, int dataRow) {
        if (reader.FieldCount != expected.Length)
            throw new InvalidDataException($"The SST memory row width differs at worksheet row {dataRow + 1}.");
        for (int column = 0; column < expected.Length; column++) {
            bool valid;
            if (expected[column] is string text && policy == TextPolicy.Utf8) {
#if OFFICEIMO_BENCHMARK_NEW_APIS
                byte[] ownedExpected = Encoding.UTF8.GetBytes(text);
                valid = reader.TryGetUtf8Text(column, out ReadOnlySpan<byte> borrowed)
                    && borrowed.SequenceEqual(ownedExpected);
#else
                throw new InvalidOperationException("Utf8 memory evidence requires current APIs.");
#endif
            } else {
                object value = reader.GetValue(column);
                valid = Equals(value, expected[column]);
                if (expected[column] is double number) valid &= reader.GetDouble(column) == number;
                else if (expected[column] is DateTime date) valid &= reader.GetDateTime(column) == date;
            }
            if (!valid)
                throw new InvalidDataException($"The SST memory value/type differs at worksheet row {dataRow + 1}, column {column + 1}.");
        }
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void ValidateSchema(ExcelWorkbookDataReader reader) {
        using DataTable schema = reader.GetSchemaTable() ?? throw new InvalidDataException("The SST reader omitted its raw schema.");
        if (schema.Rows.Count != StringHeavyWorkbookGenerator.Headers.Length)
            throw new InvalidDataException("The SST raw schema width differs.");
        for (int column = 0; column < schema.Rows.Count; column++) {
            string name = $"Column{column + 1}";
            DataRow field = schema.Rows[column];
            if (reader.GetName(column) != name || reader.GetOrdinal(name) != column || reader.GetFieldType(column) != typeof(object)
                || (string)field[SchemaTableColumn.ColumnName] != name || (int)field[SchemaTableColumn.ColumnOrdinal] != column
                || (Type)field[SchemaTableColumn.DataType] != typeof(object))
                throw new InvalidDataException($"The SST raw schema differs at column {column + 1}.");
        }
    }

    private static MemorySnapshot Snapshot(Process process, string stage, long? baseline = null) {
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);
        long managedBytes = GC.GetTotalMemory(forceFullCollection: false);
        process.Refresh();
        long peakBytes = process.PeakWorkingSet64;
        return new MemorySnapshot(stage, managedBytes, managedBytes - (baseline ?? managedBytes),
            process.WorkingSet64, peakBytes > 0 ? peakBytes : null);
    }

    private sealed record Fixture(byte[] Bytes, string SourceSha256, string Sha256, bool FromSavedInput);
    private sealed record MemorySnapshot(string Stage, long ManagedRetainedBytes, long ManagedDeltaFromBeforeOpenBytes,
        long ProcessWorkingSetBytes, long? ProcessPeakWorkingSetBytes);
}
