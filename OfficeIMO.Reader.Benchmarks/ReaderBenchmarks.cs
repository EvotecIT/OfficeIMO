using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using OfficeIMO.Markdown;
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.Markdown;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Reader.Word;
using OfficeIMO.Reader.Xml;
using OfficeIMO.Reader.Yaml;
using OfficeIMO.Reader.Zip;

namespace OfficeIMO.Reader.Benchmarks;

[MemoryDiagnoser]
[ShortRunJob(RuntimeMoniker.Net80)]
public class ReaderDocumentBenchmarks {
    private OfficeDocumentReader _reader = null!;
    private ReaderBenchmarkInput _input = null!;
    private ReaderOptions _options = null!;

    [ParamsSource(nameof(FormatNames))]
    public string Format { get; set; } = string.Empty;

    public IEnumerable<string> FormatNames() => ReaderBenchmarkCorpus.Names;

    [GlobalSetup]
    public void Setup() {
        _reader = CreateReader();
        _input = ReaderBenchmarkCorpus.Get(Format);
        _options = new ReaderOptions {
            ComputeHashes = false,
            MaxChars = 4_000,
            MaxTableRows = 5_000
        };
    }

    [Benchmark]
    public OfficeDocumentReadResult ReadDocument() =>
        _reader.ReadDocument(_input.Bytes, _input.SourceName, _options);

    internal static OfficeDocumentReader CreateReader() => new OfficeDocumentReaderBuilder()
        .AddCsvHandler()
        .AddEmailHandlers()
        .AddEpubHandler()
        .AddExcelHandler()
        .AddHtmlHandler()
        .AddJsonHandler()
        .AddMarkdownHandler()
        .AddPdfHandler()
        .AddPowerPointHandler()
        .AddRtfHandler()
        .AddVisioHandler()
        .AddWordHandler()
        .AddXmlHandler()
        .AddYamlHandler()
        .AddZipHandler()
        .Build();
}

[MemoryDiagnoser]
[ShortRunJob(RuntimeMoniker.Net80)]
public class ReaderDetectionBenchmarks {
    private ReaderBenchmarkInput _input = null!;
    private ReaderDetectionOptions _options = null!;

    [Params("BinaryPowerPoint", "PowerPoint", "Markdown", "Pdf", "Word", "Epub", "Visio")]
    public string Format { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup() {
        _input = ReaderBenchmarkCorpus.Get(Format);
        _options = new ReaderDetectionOptions { Mode = ReaderDetectionMode.PreferContent };
    }

    [Benchmark]
    public ReaderDetectionResult Detect() =>
        OfficeDocumentReader.Default.Detect(_input.Bytes, _input.SourceName, _options);
}

[MemoryDiagnoser]
[ShortRunJob(RuntimeMoniker.Net80)]
public class ReaderTransportBenchmarks {
    private OfficeDocumentReadResult _result = null!;
    private string _json = string.Empty;

    [GlobalSetup]
    public void Setup() {
        ReaderBenchmarkInput input = ReaderBenchmarkCorpus.Get("Markdown");
        _result = ReaderDocumentBenchmarks.CreateReader().ReadDocument(
            input.Bytes,
            input.SourceName,
            new ReaderOptions { ComputeHashes = false });
        _json = _result.ToJson();
    }

    [Benchmark]
    public string Serialize() => _result.ToJson();

    [Benchmark]
    public OfficeDocumentReadResult Deserialize() => OfficeDocumentReadResultJson.Deserialize(_json);
}

[MemoryDiagnoser]
[ShortRunJob(RuntimeMoniker.Net80)]
public class ReaderHierarchicalChunkingBenchmarks {
    private OfficeDocumentReadResult _document = null!;
    private ReaderHierarchicalChunkingOptions _options = null!;
    private ReaderChunkHierarchyResult _hierarchy = null!;

    [GlobalSetup]
    public void Setup() {
        ReaderBenchmarkInput input = ReaderBenchmarkCorpus.Get("Markdown");
        _document = ReaderDocumentBenchmarks.CreateReader().ReadDocument(
            input.Bytes,
            input.SourceName,
            new ReaderOptions { ComputeHashes = false, MaxChars = 4_000, MaxTableRows = 5_000 });
        _options = new ReaderHierarchicalChunkingOptions {
            MaxTokens = 512,
            OverlapTokens = 64,
            MaxInputChunks = 10_000,
            MaxOutputChunks = 50_000,
            IncludeContextInText = true
        };
        _hierarchy = ReaderHierarchicalChunker.Chunk(_document, _options);
    }

    [Benchmark]
    public ReaderChunkHierarchyResult ChunkHierarchy() =>
        ReaderHierarchicalChunker.Chunk(_document, _options);

    [Benchmark]
    public string SerializeHierarchy() => _hierarchy.ToJson();
}

[MemoryDiagnoser]
[ShortRunJob(RuntimeMoniker.Net80)]
public class ReaderMarkdownPipelineBenchmarks {
    private byte[] _bytes = Array.Empty<byte>();
    private string _markdown = string.Empty;
    private string _markdownWithoutTables = string.Empty;
    private OfficeDocumentReader _headingReader = null!;
    private OfficeDocumentReader _paragraphReader = null!;
    private ReaderOptions _readerOptions = null!;

    [GlobalSetup]
    public void Setup() {
        ReaderBenchmarkInput input = ReaderBenchmarkCorpus.Get("Markdown");
        _bytes = input.Bytes;
        _markdown = System.Text.Encoding.UTF8.GetString(input.Bytes);
        _markdownWithoutTables = string.Join(
            "\n",
            _markdown.Split('\n').Where(static line => !line.StartsWith("|", StringComparison.Ordinal)));
        _readerOptions = new ReaderOptions {
            ComputeHashes = false,
            MaxChars = 4_000,
            MaxTableRows = 5_000
        };
        _headingReader = new OfficeDocumentReaderBuilder().AddMarkdownHandler().Build();
        _paragraphReader = new OfficeDocumentReaderBuilder()
            .AddMarkdownHandler(new ReaderMarkdownOptions { ChunkByHeadings = false })
            .Build();
    }

    [Benchmark]
    public MarkdownParseResult ParseWithSyntaxTreeAndTables() => MarkdownReader.ParseWithSyntaxTree(_markdown);

    [Benchmark]
    public MarkdownParseResult ParseWithSyntaxTreeWithoutTables() => MarkdownReader.ParseWithSyntaxTree(_markdownWithoutTables);

    [Benchmark]
    public ReaderChunk[] ReadHeadingAndTableChunks() =>
        _headingReader.Read(_bytes, "handbook.md", _readerOptions).ToArray();

    [Benchmark]
    public ReaderChunk[] ReadParagraphChunks() =>
        _paragraphReader.Read(_bytes, "handbook.md", _readerOptions).ToArray();
}
