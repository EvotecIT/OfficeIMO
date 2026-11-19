#if NET10_0_OR_GREATER
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

public partial class ExcelDataReaderWriteBenchmarks {
    [GlobalSetup(Target = nameof(OneDocument))]
    public void SetupOneDocument() {
        Initialize();
        ValidateOutput(nameof(OneDocument), ExcelLibraryComparisonRunner.OneDocumentWriteDataReaderPlainBytes(_table));
    }

    /// <summary>Streams the same reader through OneDocument's forward-only XLSX writer.</summary>
    [Benchmark]
    [BenchmarkCategory("library:OneDocument.Spreadsheets", "case:Default")]
    public int OneDocument()
        => ExcelLibraryComparisonRunner.OneDocumentWriteDataReaderPlainBytes(_table).Length;
}
#endif
