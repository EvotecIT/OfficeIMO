using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>Publishable paragraph-creation comparison against the MIT-licensed Open XML SDK.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "PublicEvidence", "Create")]
public class WordParagraphOpenXmlEvidenceBenchmarks {
    private WordCreateParagraphComparisonBenchmarks _workload = null!;

    [Params(100, 1000)]
    public int ItemCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _workload = new WordCreateParagraphComparisonBenchmarks { ItemCount = ItemCount };
        WordBenchmarkCorpus.ValidateParagraphDocument(_workload.OfficeIMO(), ItemCount);
        WordBenchmarkCorpus.ValidateParagraphDocument(OpenXmlSdk(), ItemCount);
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => _workload.OfficeIMO();

    [Benchmark]
    public byte[] OpenXmlSdk() => WordOpenXmlEvidenceWorkloads.CreateParagraphs(ItemCount);
}

/// <summary>Publishable structured-report comparison against the MIT-licensed Open XML SDK.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "PublicEvidence", "StructuredReport")]
public class WordReportOpenXmlEvidenceBenchmarks {
    private WordCreateReportComparisonBenchmarks _workload = null!;

    [Params(100, 1000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _workload = new WordCreateReportComparisonBenchmarks { RowCount = RowCount };
        WordBenchmarkCorpus.ValidateReportDocument(
            _workload.OfficeIMO(), RowCount, requireOfficeCompatibleDefaults: true);
        WordBenchmarkCorpus.ValidateReportDocument(
            OpenXmlSdk(), RowCount, requireOfficeCompatibleDefaults: true);
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => _workload.OfficeIMO();

    [Benchmark]
    public byte[] OpenXmlSdk() => WordOpenXmlEvidenceWorkloads.CreateReport(RowCount);
}

/// <summary>Publishable full-read comparison against the MIT-licensed Open XML SDK.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "PublicEvidence", "Read")]
public class WordReadOpenXmlEvidenceBenchmarks {
    private WordReadComparisonBenchmarks _workload = null!;

    [Params(100, 1000)]
    public int ItemCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _workload = new WordReadComparisonBenchmarks { ItemCount = ItemCount };
        _workload.SetupOfficeAndOpenXml();
    }

    [Benchmark(Baseline = true)]
    public WordReadObservation OfficeIMO() => _workload.OfficeIMO();

    [Benchmark]
    public WordReadObservation OpenXmlSdk() => _workload.OpenXmlSdk();
}

/// <summary>Publishable load-replace-save comparison against the MIT-licensed Open XML SDK.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "PublicEvidence", "Replace")]
public class WordReplaceOpenXmlEvidenceBenchmarks {
    private WordRichReplaceEvidenceWorkload _workload = null!;

    [Params(100, 1000)]
    public int ItemCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _workload = new WordRichReplaceEvidenceWorkload(ItemCount);
        int expectedStyles = WordBenchmarkCorpus.CountStyleDefinitions(_workload.Fixture);
        byte[] office = _workload.OfficeIMO();
        byte[] sdk = _workload.OpenXmlSdk();
        WordBenchmarkCorpus.ValidateReplacedDocument(office, ItemCount);
        WordBenchmarkCorpus.ValidateReplacedDocument(sdk, ItemCount);
        if (WordBenchmarkCorpus.CountStyleDefinitions(office) != expectedStyles ||
            WordBenchmarkCorpus.CountStyleDefinitions(sdk) != expectedStyles) {
            throw new InvalidDataException("A replacement implementation changed the rich input style catalog.");
        }
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => _workload.OfficeIMO();

    [Benchmark]
    public byte[] OpenXmlSdk() => _workload.OpenXmlSdk();
}

internal static class WordOpenXmlEvidenceValidation {
    internal static void RunAll() {
        foreach (int itemCount in new[] { 100, 1000 }) {
            new WordParagraphOpenXmlEvidenceBenchmarks { ItemCount = itemCount }.Setup();
            new WordReportOpenXmlEvidenceBenchmarks { RowCount = itemCount }.Setup();
            new WordReadOpenXmlEvidenceBenchmarks { ItemCount = itemCount }.Setup();
            new WordReplaceOpenXmlEvidenceBenchmarks { ItemCount = itemCount }.Setup();
        }
    }
}
