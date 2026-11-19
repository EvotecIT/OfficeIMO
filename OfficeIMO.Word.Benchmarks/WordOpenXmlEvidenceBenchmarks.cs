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
        string detail = $"ItemCount={ItemCount}; {ItemCount} body paragraphs with the expected text; passes the Open XML SDK validator";
        byte[] officeImo = _workload.OfficeIMO();
        WordBenchmarkCorpus.Check("OfficeIMO", "Default", detail,
            () => WordBenchmarkCorpus.ValidateParagraphDocument(officeImo, ItemCount));
        byte[] openXmlSdk = OpenXmlSdk();
        WordBenchmarkCorpus.Check("Open XML SDK", "Default", detail,
            () => WordBenchmarkCorpus.ValidateParagraphDocument(openXmlSdk, ItemCount));
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() => _workload.OfficeIMO();

    [Benchmark]
    [BenchmarkCategory("library:Open XML SDK", "case:Default")]
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
        string detail = $"RowCount={RowCount}; title, summary, {RowCount + 1}-row table with Office-compatible defaults, header and footer with the expected text; passes the Open XML SDK validator";
        byte[] officeImo = _workload.OfficeIMO();
        WordBenchmarkCorpus.Check("OfficeIMO", "Default", detail,
            () => WordBenchmarkCorpus.ValidateReportDocument(officeImo, RowCount, requireOfficeCompatibleDefaults: true));
        byte[] openXmlSdk = OpenXmlSdk();
        WordBenchmarkCorpus.Check("Open XML SDK", "Default", detail,
            () => WordBenchmarkCorpus.ValidateReportDocument(openXmlSdk, RowCount, requireOfficeCompatibleDefaults: true));
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() => _workload.OfficeIMO();

    [Benchmark]
    [BenchmarkCategory("library:Open XML SDK", "case:Default")]
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
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public WordReadObservation OfficeIMO() => _workload.OfficeIMO();

    [Benchmark]
    [BenchmarkCategory("library:Open XML SDK", "case:Default")]
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
        string detail = $"ItemCount={ItemCount}; {ItemCount} paragraphs with the placeholder replaced, {expectedStyles} style definitions kept; passes the Open XML SDK validator";
        byte[] office = _workload.OfficeIMO();
        WordBenchmarkCorpus.Check("OfficeIMO", "Default", detail, () => ValidateReplaced(office, expectedStyles));
        byte[] sdk = _workload.OpenXmlSdk();
        WordBenchmarkCorpus.Check("Open XML SDK", "Default", detail, () => ValidateReplaced(sdk, expectedStyles));
    }

    private void ValidateReplaced(byte[] payload, int expectedStyles) {
        WordBenchmarkCorpus.ValidateReplacedDocument(payload, ItemCount);
        if (WordBenchmarkCorpus.CountStyleDefinitions(payload) != expectedStyles) {
            throw new InvalidDataException("The replacement changed the rich input style catalog.");
        }
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() => _workload.OfficeIMO();

    [Benchmark]
    [BenchmarkCategory("library:Open XML SDK", "case:Default")]
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
