using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Which PDFsharp-compatible engine produced the inputs for page operations.</summary>
public enum PdfSharpEngineSource {
    OfficeIMO,
    MigraDoc
}

internal static class PdfSharpEngineSources {
    internal static byte[] Generate(PdfSharpEngineSource source, PdfBenchmarkScenario scenario) => source switch {
        PdfSharpEngineSource.OfficeIMO => OfficeImoPdfGenerator.Generate(scenario),
        PdfSharpEngineSource.MigraDoc => MigraDocPdfGenerator.Generate(scenario),
        _ => throw new ArgumentOutOfRangeException(nameof(source))
    };
}

/// <summary>
/// The comparison project's structured-report generation lane, with MigraDoc
/// rendered by the engine this build selected (FreePDF or PDFsharp).
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfSharpEngineGenerationBenchmarks {
    private PdfBenchmarkScenario _scenario = null!;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [GlobalSetup]
    public void Setup() {
        _scenario = PdfBenchmarkScenario.Get(Scale);
        PdfBenchmarkValidation.ValidateGenerated(OfficeIMO(), _scenario, nameof(OfficeIMO));
        PdfBenchmarkValidation.ValidateGenerated(MigraDoc(), _scenario, nameof(MigraDoc));
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => OfficeImoPdfGenerator.Generate(_scenario);

    [Benchmark]
    public byte[] MigraDoc() => MigraDocPdfGenerator.Generate(_scenario);
}

/// <summary>Merges many identical-shape source PDFs into one ordered document.</summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfSharpEngineMergeBenchmarks {
    private byte[][] _sources = null!;
    private IReadOnlyList<IReadOnlyList<PdfExpectedPage>> _expectedPages = null!;

    [Params(PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [Params(PdfSharpEngineSource.OfficeIMO, PdfSharpEngineSource.MigraDoc)]
    public PdfSharpEngineSource Producer { get; set; }

    [GlobalSetup]
    public void Setup() {
        PdfManipulationScenario scenario = PdfManipulationScenario.Get(Scale);
        IReadOnlyList<(PdfBenchmarkScenario Scenario, int[] Pages)> documents =
            scenario.ExpectedMergeDocuments(PdfManipulationDensity.Standard);
        _sources = documents.Select(item => PdfSharpEngineSources.Generate(Producer, item.Scenario)).ToArray();
        for (int index = 0; index < _sources.Length; index++) {
            PdfBenchmarkValidation.ValidateGenerated(_sources[index], documents[index].Scenario, Producer + " source " + (index + 1));
        }

        _expectedPages = new[] {
            (IReadOnlyList<PdfExpectedPage>)documents
                .SelectMany(item => item.Pages.Select(page => PdfBenchmarkValidation.ExpectedPage(item.Scenario, page)))
                .ToArray()
        };
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(PdfSharp), PdfSharp());
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => PdfManipulationEngines.MergeWithOfficeImo(_sources);

    [Benchmark]
    public byte[] PdfSharp() => PdfManipulationEngines.MergeWithPdfSharp(_sources);

    private void Validate(string engine, byte[] output) =>
        PdfManipulationValidation.Validate(new[] { output }, _expectedPages, engine + " merging " + Producer);
}

/// <summary>Splits identical PDF bytes into single pages or consecutive page bundles.</summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfSharpEngineSplitBenchmarks {
    private byte[] _source = null!;
    private int _pagesPerDocument;
    private IReadOnlyList<IReadOnlyList<PdfExpectedPage>> _expectedPages = null!;

    [Params(PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [Params(PdfSharpEngineSource.OfficeIMO, PdfSharpEngineSource.MigraDoc)]
    public PdfSharpEngineSource Producer { get; set; }

    [Params(PdfSplitWorkflow.EveryPage, PdfSplitWorkflow.Bundles)]
    public PdfSplitWorkflow Workflow { get; set; }

    [GlobalSetup]
    public void Setup() {
        PdfManipulationScenario scenario = PdfManipulationScenario.Get(Scale);
        PdfBenchmarkScenario sourceScenario = scenario.SourceDocument(PdfManipulationDensity.Standard);
        _source = PdfSharpEngineSources.Generate(Producer, sourceScenario);
        PdfBenchmarkValidation.ValidateGenerated(_source, sourceScenario, Producer.ToString());
        _pagesPerDocument = Workflow == PdfSplitWorkflow.EveryPage ? 1 : scenario.PagesPerBundle;
        _expectedPages = scenario.ExpectedSplitPages(Workflow)
            .Select(pages => (IReadOnlyList<PdfExpectedPage>)pages
                .Select(page => PdfBenchmarkValidation.ExpectedPage(sourceScenario, page))
                .ToArray())
            .ToArray();
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(PdfSharp), PdfSharp());
    }

    [Benchmark(Baseline = true)]
    public byte[][] OfficeIMO() => PdfManipulationEngines.SplitWithOfficeImo(_source, _pagesPerDocument);

    [Benchmark]
    public byte[][] PdfSharp() => PdfManipulationEngines.SplitWithPdfSharp(_source, _pagesPerDocument);

    private void Validate(string engine, byte[][] outputs) =>
        PdfManipulationValidation.Validate(outputs, _expectedPages, engine + " splitting " + Producer);
}

/// <summary>Extracts a deterministic non-contiguous page selection in reverse order.</summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfSharpEnginePageSelectionBenchmarks {
    private byte[] _source = null!;
    private int[] _pageNumbers = null!;
    private IReadOnlyList<IReadOnlyList<PdfExpectedPage>> _expectedPages = null!;

    [Params(PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [Params(PdfSharpEngineSource.OfficeIMO, PdfSharpEngineSource.MigraDoc)]
    public PdfSharpEngineSource Producer { get; set; }

    [GlobalSetup]
    public void Setup() {
        PdfManipulationScenario scenario = PdfManipulationScenario.Get(Scale);
        PdfBenchmarkScenario sourceScenario = scenario.SourceDocument(PdfManipulationDensity.Standard);
        _source = PdfSharpEngineSources.Generate(Producer, sourceScenario);
        PdfBenchmarkValidation.ValidateGenerated(_source, sourceScenario, Producer.ToString());
        _pageNumbers = scenario.SelectedPages();
        _expectedPages = new[] {
            (IReadOnlyList<PdfExpectedPage>)_pageNumbers
                .Select(page => PdfBenchmarkValidation.ExpectedPage(sourceScenario, page))
                .ToArray()
        };
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(PdfSharp), PdfSharp());
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => PdfManipulationEngines.SelectWithOfficeImo(_source, _pageNumbers);

    [Benchmark]
    public byte[] PdfSharp() => PdfManipulationEngines.SelectWithPdfSharp(_source, _pageNumbers);

    private void Validate(string engine, byte[] output) =>
        PdfManipulationValidation.Validate(new[] { output }, _expectedPages, engine + " selecting pages from " + Producer);
}
