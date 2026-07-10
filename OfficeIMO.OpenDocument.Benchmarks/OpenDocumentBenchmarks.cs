using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;

namespace OfficeIMO.OpenDocument.Benchmarks;

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class OpenDocumentPackageBenchmarks {
    private byte[] _odt = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        using OdtDocument document = OdtDocument.Create();
        for (int index = 0; index < 2_000; index++) document.AddParagraph("Paragraph " + index + " with structured text and whitespace.");
        _odt = document.ToBytes();
    }

    [Benchmark]
    public int OpenAndEnumerateOdt() {
        using OdtDocument document = OdtDocument.Open(new MemoryStream(_odt, writable: false));
        return document.ContentBlocks.Count;
    }

    [Benchmark]
    public int CreateSparseOdsAtExtremeCoordinate() {
        using OdsDocument document = OdsDocument.Create();
        document.AddSheet("Sparse").Cell(1_000_000, 500_000).SetString("edge");
        return document.ToBytes().Length;
    }
}

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class OpenDocumentFormulaBenchmarks {
    private OdsDocument _document = null!;

    [GlobalSetup]
    public void Setup() {
        _document = OdsDocument.Create();
        OdsSheet sheet = _document.AddSheet("Data");
        for (int row = 0; row < 1_000; row++) sheet.Cell(row, 0).SetNumber(row + 1D);
        sheet.Cell(0, 1).Formula = "of:=SUM([.A1:.A1000])+AVERAGE([.A1:.A1000])";
    }

    [GlobalCleanup]
    public void Cleanup() => _document.Dispose();

    [Benchmark]
    public double EvaluateRangeFormula() => OdsFormulaEvaluator.EvaluateCell(_document, "Data", 0, 1).Value.AsNumber();
}
