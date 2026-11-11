using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

public partial class MarkPflug65KXlsxBenchmarks {
    /// <summary>Reads the same decimal observations while retaining default double-valued cells.</summary>
    [Benchmark]
    public ExcelReadObservation OfficeIMO_DefaultNumbers() {
        using var reader = ExcelDocument.OpenReader(MarkPflug65KFixture.XlsxPath);
        return Validate(nameof(OfficeIMO_DefaultNumbers), Observe(reader));
    }

    /// <summary>Combines default double-valued cells with worksheet prefetch.</summary>
    [Benchmark]
    public ExcelReadObservation OfficeIMO_DefaultNumbers_Prefetch() {
        using var reader = ExcelDocument.OpenReader(MarkPflug65KFixture.XlsxPath,
            new ExcelReaderOptions { EnableWorksheetPrefetch = true });
        return Validate(nameof(OfficeIMO_DefaultNumbers_Prefetch), Observe(reader));
    }

    /// <summary>Reads the same values through a loaded document's snapshot route.</summary>
    [Benchmark]
    public ExcelReadObservation OfficeIMO_Loaded() {
        using var document = ExcelDocument.Load(new MemoryStream(_inputBytes, writable: false));
        using var reader = document.CreateDataReader(new ExcelReaderOptions { NumericAsDecimal = true });
        return Validate(nameof(OfficeIMO_Loaded), Observe(reader));
    }

    /// <summary>Prototype single-pass delivery: rows are read while inflation and validation continue.</summary>
    [Benchmark]
    public ExcelReadObservation OfficeIMO_SinglePass() {
        using var reader = ExcelDocument.OpenReader(MarkPflug65KFixture.XlsxPath,
            new ExcelReaderOptions { NumericAsDecimal = true, SheetName = null, EnableSinglePassReading = true });
        return Validate(nameof(OfficeIMO_SinglePass), Observe(reader));
    }

    /// <summary>Excludes file access using the same immutable package as the memory comparison.</summary>
    [Benchmark]
    public ExcelReadObservation OfficeIMO_Memory() {
        using var stream = new MemoryStream(_inputBytes, writable: false);
        using var reader = ExcelDocument.OpenReader(stream, new ExcelReaderOptions { NumericAsDecimal = true });
        return Validate(nameof(OfficeIMO_Memory), Observe(reader));
    }

#if NET10_0_OR_GREATER
    /// <summary>Uses the published direct in-memory factory, including its default checksum validation.</summary>
    [Benchmark]
    public ExcelReadObservation ExcelReaderNet_Memory() {
        using var reader = global::ExcelReader.Core.Reader.Excel.FromXlsx(_inputBytes.AsMemory());
        return Validate(nameof(ExcelReaderNet_Memory), ObserveExcelReader(reader));
    }

    /// <summary>Measures the opt-in string policy without changing the values observed.</summary>
    [Benchmark]
    public ExcelReadObservation ExcelReaderNet_InternStrings() {
        var options = global::ExcelReader.Core.Reader.ExcelReaderOptions.Default with { InternStrings = true };
        using var reader = global::ExcelReader.Core.Reader.Excel.FromXlsxFile(MarkPflug65KFixture.XlsxPath, options);
        return Validate(nameof(ExcelReaderNet_InternStrings), ObserveExcelReader(reader));
    }
#endif
}

#if NET10_0_OR_GREATER
public partial class MarkPflug65KXlsbBenchmarks {
    /// <summary>Enables background decompression while retaining every other published reader default.</summary>
    [Benchmark]
    public ExcelReadObservation ExcelReaderNet_Prefetch() {
        var options = global::ExcelReader.Core.Reader.ExcelReaderOptions.Default with {
            PrefetchDecompression = true
        };
        using var reader = global::ExcelReader.Core.Reader.Excel.FromXlsbFile(
            MarkPflug65KFixture.XlsbPath, options);
        return Validate(nameof(ExcelReaderNet_Prefetch), MarkPflug65KXlsxBenchmarks.ObserveExcelReader(reader));
    }
}
#endif
