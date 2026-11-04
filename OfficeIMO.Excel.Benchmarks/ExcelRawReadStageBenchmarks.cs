#if NET10_0_OR_GREATER
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Locates preparation versus delivery costs using the validated generated XLSX fixture.</summary>
[MemoryDiagnoser]
[WarmupCount(8), IterationCount(32)]
public class ExcelRawReadStageBenchmarks {
    private readonly ExcelReaderGeneratedBenchmarks _fixture = new() { Format = ExcelFileFormat.Xlsx, RowCount = 50_000 };
    private string _path = string.Empty;
    private const int BatchSize = 32;
    private readonly ExcelWorkbookDataReader?[] _prepared = new ExcelWorkbookDataReader?[BatchSize];

    [GlobalSetup]
    public void Setup() {
        _fixture.Setup(); // Full independent-reader/value preflight, outside measurement.
        _path = _fixture.InputPath;
    }

    [GlobalCleanup]
    public void Cleanup() {
        CloseRows();
        _fixture.Cleanup();
    }

    [Benchmark(OperationsPerInvoke = BatchSize)]
    public bool PrepareAndDispose() {
        for (int index = 0; index < BatchSize; index++) {
            using var reader = ExcelDocument.OpenReader(_path);
            if (!reader.HasRows) throw new InvalidDataException("Missing rows.");
        }
        return true;
    }

    [IterationSetup(Target = nameof(ConsumePreparedRows))]
    public void PrepareRows() {
        try {
            for (int index = 0; index < BatchSize; index++) {
                _prepared[index] = ExcelDocument.OpenReader(_path);
                if (!_prepared[index]!.HasRows) throw new InvalidDataException("Missing rows.");
            }
        } catch {
            CloseRows();
            throw;
        }
    }

    [Benchmark(OperationsPerInvoke = BatchSize)]
    public long ConsumePreparedRows() {
        long checksum = 0;
        foreach (var reader in _prepared) checksum ^= _fixture.Read(reader!);
        return checksum;
    }

    [IterationCleanup(Target = nameof(ConsumePreparedRows))]
    public void CloseRows() {
        for (int index = 0; index < BatchSize; index++) {
            _prepared[index]?.Dispose();
            _prepared[index] = null;
        }
    }
}
#endif
