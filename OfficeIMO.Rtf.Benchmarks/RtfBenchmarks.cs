using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Columns;
using OfficeIMO.Html;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Rtf.Markdown;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;

namespace OfficeIMO.Rtf.Benchmarks;

[MemoryDiagnoser]
[OperationsPerSecond]
public class RtfCoreBenchmarks {
    private string _rtf = string.Empty;

    [ParamsSource(nameof(CorpusScales))]
    public string Scale { get; set; } = string.Empty;

    public IEnumerable<string> CorpusScales() => RtfBenchmarkCorpus.Scales;

    [GlobalSetup]
    public void Setup() => _rtf = RtfBenchmarkCorpus.Get(Scale).Rtf;

    [Benchmark]
    [BenchmarkCategory("Parse")]
    public RtfReadResult Parse() => RtfDocument.Read(_rtf);

    [Benchmark]
    [BenchmarkCategory("Lossless")]
    public string LosslessRoundTrip() => RtfDocument.Read(_rtf).EditLossless().ToRtf();

    [Benchmark]
    [BenchmarkCategory("SemanticWrite")]
    public string SemanticRewrite() => RtfDocument.Read(_rtf).Document.ToRtf();
}

[MemoryDiagnoser]
[OperationsPerSecond]
public class RtfAdapterBenchmarks {
    private RtfDocument _document = null!;
    private RtfPdfSaveOptions _pdfOptions = null!;

    [ParamsSource(nameof(CorpusScales))]
    public string Scale { get; set; } = string.Empty;

    public IEnumerable<string> CorpusScales() => RtfBenchmarkCorpus.Scales;

    [GlobalSetup]
    public void Setup() {
        _document = RtfDocument.Read(RtfBenchmarkCorpus.Get(Scale).Rtf).Document;
        _pdfOptions = RtfBenchmarkSupport.CreatePdfSaveOptions();
    }

    [Benchmark]
    [BenchmarkCategory("Html")]
    public string Html() => _document.ToHtml();

    [Benchmark]
    [BenchmarkCategory("Markdown")]
    public string Markdown() => _document.ToMarkdown();

    [Benchmark]
    [BenchmarkCategory("Pdf")]
    public byte[] Pdf() => _document.ToPdfDocument(_pdfOptions).ToBytes();

    [Benchmark]
    [BenchmarkCategory("Word")]
    public int WordModel() {
        using WordDocument word = _document.ToWordDocument();
        return _document.Blocks.Count;
    }

    [Benchmark]
    [BenchmarkCategory("Word")]
    public byte[] Word() {
        using WordDocument word = _document.ToWordDocument();
        using MemoryStream stream = word.ToDocxStream();
        return stream.ToArray();
    }

    [Benchmark]
    [BenchmarkCategory("Reader")]
    public ReaderChunk[] Reader() => DocumentReaderRtfExtensions.ReadRtfDocument(_document).ToArray();
}
