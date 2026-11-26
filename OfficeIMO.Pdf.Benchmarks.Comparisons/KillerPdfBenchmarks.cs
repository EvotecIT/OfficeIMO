#if NET10_0_OR_GREATER
using System.Text;
using BenchmarkDotNet.Attributes;
using KillerPdf.Engine.Documents;
using KillerPdf.Engine.Editing;
using KillerPdf.Engine.Fonts;
using KillerPdf.Engine.Rendering;
using OfficeIMO.Drawing;
using KillerPdfDocument = KillerPdf.Engine.Documents.PdfDocument;
using OfficePdfDocument = OfficeIMO.Pdf.PdfDocument;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Which library produced the input PDF.</summary>
public enum KillerPdfInputProducer {
    OfficeIMO,
    MigraDoc
}

internal static class KillerPdfEngines {
    internal static byte[] Generate(KillerPdfInputProducer producer, PdfBenchmarkScenario scenario) => producer switch {
        KillerPdfInputProducer.OfficeIMO => OfficeImoPdfGenerator.Generate(scenario),
        KillerPdfInputProducer.MigraDoc => MigraDocPdfGenerator.Generate(scenario),
        _ => throw new ArgumentOutOfRangeException(nameof(producer))
    };

    internal static PdfReadObservation ReadWithOfficeImo(byte[] pdf) {
        using var input = new MemoryStream(pdf, writable: false);
        using OfficePdfDocument document = OfficePdfDocument.Load(input);
        string text = document.ExtractText();
        return PdfBenchmarkValidation.Observe(document.Pages.Count, text);
    }

    internal static PdfReadObservation ReadWithKillerPdf(byte[] pdf) {
        KillerPdfDocument document = KillerPdfDocument.Open(pdf);
        var reader = new PdfPageContentReader(document);
        var text = new StringBuilder();
        for (int page = 0; page < reader.PageCount; page++) {
            text.Append(reader.Read(page).Text);
            text.Append('\n');
        }

        return PdfBenchmarkValidation.Observe(reader.PageCount, text.ToString());
    }

    internal static byte[] MergeWithKillerPdf(byte[][] sources) {
        var editor = new PdfIncrementalPageEditor(KillerPdfDocument.Open(sources[0]));
        for (int index = 1; index < sources.Length; index++) {
            editor.AddImportedDocument(KillerPdfDocument.Open(sources[index]));
        }

        return editor.Build();
    }

    internal static byte[] SelectWithKillerPdf(byte[] source, int[] pageNumbers) {
        KillerPdfDocument document = KillerPdfDocument.Open(source);
        return PdfDocumentSplitter.Extract(document, pageNumbers.Select(static page => page - 1).ToArray());
    }

    internal static byte[][] SplitWithKillerPdf(byte[] source, int pagesPerDocument) {
        KillerPdfDocument document = KillerPdfDocument.Open(source);
        IReadOnlyList<PdfSplitOutput> parts = PdfDocumentSplitter.Split(
            document,
            new PdfSplitOptions { Mode = PdfSplitMode.PageCount, PagesPerPart = pagesPerDocument });
        var outputs = new byte[parts.Count][];
        for (int index = 0; index < parts.Count; index++) {
            outputs[index] = parts[index].Document.ToArray();
        }

        return outputs;
    }
}

/// <summary>Opens identical bytes, enumerates every page and extracts the complete text payload.</summary>
[MemoryDiagnoser]
[RankColumn]
public class KillerPdfReadBenchmarks {
    private byte[] _pdf = null!;
    private PdfBenchmarkScenario _scenario = null!;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [Params(KillerPdfInputProducer.OfficeIMO, KillerPdfInputProducer.MigraDoc)]
    public KillerPdfInputProducer Producer { get; set; }

    [GlobalSetup]
    public void Setup() {
        _scenario = PdfBenchmarkScenario.Get(Scale);
        _pdf = KillerPdfEngines.Generate(Producer, _scenario);
        PdfBenchmarkValidation.ValidateGenerated(_pdf, _scenario, Producer.ToString());
        PdfBenchmarkValidation.ValidateRead(OfficeIMO(), _scenario, "OfficeIMO reading " + Producer);
        PdfBenchmarkValidation.ValidateRead(KillerPdf(), _scenario, "KillerPDF reading " + Producer);
    }

    [Benchmark(Baseline = true)]
    public PdfReadObservation OfficeIMO() => KillerPdfEngines.ReadWithOfficeImo(_pdf);

    [Benchmark]
    public PdfReadObservation KillerPdf() => KillerPdfEngines.ReadWithKillerPdf(_pdf);
}

/// <summary>Merges many identical-shape source PDFs into one ordered document.</summary>
[MemoryDiagnoser]
[RankColumn]
public class KillerPdfMergeBenchmarks {
    private byte[][] _sources = null!;
    private IReadOnlyList<IReadOnlyList<PdfExpectedPage>> _expectedPages = null!;

    [Params(PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [Params(KillerPdfInputProducer.OfficeIMO, KillerPdfInputProducer.MigraDoc)]
    public KillerPdfInputProducer Producer { get; set; }

    [GlobalSetup]
    public void Setup() {
        PdfManipulationScenario scenario = PdfManipulationScenario.Get(Scale);
        IReadOnlyList<(PdfBenchmarkScenario Scenario, int[] Pages)> documents =
            scenario.ExpectedMergeDocuments(PdfManipulationDensity.Standard);
        _sources = documents.Select(item => KillerPdfEngines.Generate(Producer, item.Scenario)).ToArray();
        for (int index = 0; index < _sources.Length; index++) {
            PdfBenchmarkValidation.ValidateGenerated(_sources[index], documents[index].Scenario, Producer + " source " + (index + 1));
        }

        _expectedPages = new[] {
            (IReadOnlyList<PdfExpectedPage>)documents
                .SelectMany(item => item.Pages.Select(page => PdfBenchmarkValidation.ExpectedPage(item.Scenario, page)))
                .ToArray()
        };
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(KillerPdf), KillerPdf());
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => PdfManipulationEngines.MergeWithOfficeImo(_sources);

    /// <summary>KillerPDF's documented route: import every page into the first document as one incremental revision.</summary>
    [Benchmark]
    public byte[] KillerPdf() => KillerPdfEngines.MergeWithKillerPdf(_sources);

    private void Validate(string engine, byte[] output) =>
        PdfManipulationValidation.Validate(new[] { output }, _expectedPages, engine + " merging " + Producer);
}

/// <summary>Splits identical PDF bytes into single pages or consecutive page bundles.</summary>
[MemoryDiagnoser]
[RankColumn]
public class KillerPdfSplitBenchmarks {
    private byte[] _source = null!;
    private int _pagesPerDocument;
    private IReadOnlyList<IReadOnlyList<PdfExpectedPage>> _expectedPages = null!;

    [Params(PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [Params(KillerPdfInputProducer.OfficeIMO, KillerPdfInputProducer.MigraDoc)]
    public KillerPdfInputProducer Producer { get; set; }

    [Params(PdfSplitWorkflow.EveryPage, PdfSplitWorkflow.Bundles)]
    public PdfSplitWorkflow Workflow { get; set; }

    [GlobalSetup]
    public void Setup() {
        PdfManipulationScenario scenario = PdfManipulationScenario.Get(Scale);
        PdfBenchmarkScenario sourceScenario = scenario.SourceDocument(PdfManipulationDensity.Standard);
        _source = KillerPdfEngines.Generate(Producer, sourceScenario);
        PdfBenchmarkValidation.ValidateGenerated(_source, sourceScenario, Producer.ToString());
        _pagesPerDocument = Workflow == PdfSplitWorkflow.EveryPage ? 1 : scenario.PagesPerBundle;
        _expectedPages = scenario.ExpectedSplitPages(Workflow)
            .Select(pages => (IReadOnlyList<PdfExpectedPage>)pages
                .Select(page => PdfBenchmarkValidation.ExpectedPage(sourceScenario, page))
                .ToArray())
            .ToArray();
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(KillerPdf), KillerPdf());
    }

    [Benchmark(Baseline = true)]
    public byte[][] OfficeIMO() => PdfManipulationEngines.SplitWithOfficeImo(_source, _pagesPerDocument);

    [Benchmark]
    public byte[][] KillerPdf() => KillerPdfEngines.SplitWithKillerPdf(_source, _pagesPerDocument);

    private void Validate(string engine, byte[][] outputs) =>
        PdfManipulationValidation.Validate(outputs, _expectedPages, engine + " splitting " + Producer);
}

/// <summary>Extracts a deterministic non-contiguous page selection in reverse order.</summary>
[MemoryDiagnoser]
[RankColumn]
public class KillerPdfPageSelectionBenchmarks {
    private byte[] _source = null!;
    private int[] _pageNumbers = null!;
    private IReadOnlyList<IReadOnlyList<PdfExpectedPage>> _expectedPages = null!;

    [Params(PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [Params(KillerPdfInputProducer.OfficeIMO, KillerPdfInputProducer.MigraDoc)]
    public KillerPdfInputProducer Producer { get; set; }

    [GlobalSetup]
    public void Setup() {
        PdfManipulationScenario scenario = PdfManipulationScenario.Get(Scale);
        PdfBenchmarkScenario sourceScenario = scenario.SourceDocument(PdfManipulationDensity.Standard);
        _source = KillerPdfEngines.Generate(Producer, sourceScenario);
        PdfBenchmarkValidation.ValidateGenerated(_source, sourceScenario, Producer.ToString());
        _pageNumbers = scenario.SelectedPages();
        _expectedPages = new[] {
            (IReadOnlyList<PdfExpectedPage>)_pageNumbers
                .Select(page => PdfBenchmarkValidation.ExpectedPage(sourceScenario, page))
                .ToArray()
        };
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(KillerPdf), KillerPdf());
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => PdfManipulationEngines.SelectWithOfficeImo(_source, _pageNumbers);

    [Benchmark]
    public byte[] KillerPdf() => KillerPdfEngines.SelectWithKillerPdf(_source, _pageNumbers);

    private void Validate(string engine, byte[] output) =>
        PdfManipulationValidation.Validate(new[] { output }, _expectedPages, engine + " selecting pages from " + Producer);
}

/// <summary>
/// Renders the first four pages of the Medium report to PNG. KillerPDF returns BGRA pixels, so its lane
/// composites them on white and encodes them with the same OfficeIMO PNG encoder as the PDFium runner,
/// at the pixel size OfficeIMO chose for the page.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class KillerPdfRenderBenchmarks {
    private const int PageCount = 4;

    private static readonly OfficeRasterEncodingOptions EncodingOptions = new() {
        WriteResolutionMetadata = false,
        Png = new OfficePngEncodeOptions { WritePhysicalResolution = false }
    };

    private byte[] _pdf = null!;
    private (int Width, int Height)[] _sizes = null!;

    [Params(1D, 2D)]
    public double Scale { get; set; }

    [Params(KillerPdfInputProducer.OfficeIMO, KillerPdfInputProducer.MigraDoc)]
    public KillerPdfInputProducer Producer { get; set; }

    [GlobalSetup]
    public void Setup() {
        PdfBenchmarkScenario scenario = PdfBenchmarkScenario.Get(PdfBenchmarkScale.Medium);
        _pdf = KillerPdfEngines.Generate(Producer, scenario);
        PdfBenchmarkValidation.ValidateGenerated(_pdf, scenario, Producer.ToString());
        byte[][] officeImo = OfficeIMO();
        _sizes = officeImo.Select(static png => {
            OfficeRasterImage image = OfficeRasterImageDecoder.Decode(png);
            return (image.Width, image.Height);
        }).ToArray();
        byte[][] killerPdf = KillerPdf();
        for (int page = 0; page < PageCount; page++) {
            OfficeRasterImage left = OfficeRasterImageDecoder.Decode(officeImo[page]);
            OfficeRasterImage right = OfficeRasterImageDecoder.Decode(killerPdf[page]);
            if (left.Width != right.Width || left.Height != right.Height) {
                throw new InvalidDataException($"Page {page + 1}: OfficeIMO {left.Width}x{left.Height}, KillerPDF {right.Width}x{right.Height}.");
            }

            RequireInk(left.GetPixels(), $"OfficeIMO page {page + 1}");
            RequireInk(right.GetPixels(), $"KillerPDF page {page + 1}");
        }
    }

    [Benchmark(Baseline = true)]
    public byte[][] OfficeIMO() {
        using var input = new MemoryStream(_pdf, writable: false);
        using OfficePdfDocument document = OfficePdfDocument.Load(input);
        IReadOnlyList<PdfPageRenderResult> results = document.Render.Pages("1-" + PageCount, new PdfPageRenderOptions {
            Format = OfficeImageExportFormat.Png,
            Scale = Scale,
            Background = global::OfficeIMO.OfficeColor.White,
            MaxPages = PageCount,
            ContinueOnError = false
        });
        return results.Select(static result => result.Bytes ?? throw new InvalidDataException($"OfficeIMO did not render page {result.PageNumber}.")).ToArray();
    }

    [Benchmark]
    public byte[][] KillerPdf() {
        var renderer = new PdfPageRenderer(KillerPdfDocument.Open(_pdf), new EmbeddedFontsOnly());
        var pages = new byte[PageCount][];
        for (int page = 0; page < PageCount; page++) {
            (int width, int height) = _sizes[page];
            PdfRenderedPage rendered = renderer.Render(page, new PdfRenderOptions(width, height));
            ReadOnlySpan<byte> bgra = rendered.Pixels.Span;
            var rgba = new byte[bgra.Length];
            for (int index = 0; index < bgra.Length; index += 4) {
                int alpha = bgra[index + 3], inverse = 255 - alpha;
                rgba[index] = (byte)((bgra[index + 2] * alpha + 255 * inverse + 127) / 255);
                rgba[index + 1] = (byte)((bgra[index + 1] * alpha + 255 * inverse + 127) / 255);
                rgba[index + 2] = (byte)((bgra[index] * alpha + 255 * inverse + 127) / 255);
                rgba[index + 3] = 255;
            }

            pages[page] = OfficeRasterImageEncoder.Encode(
                OfficeRasterImage.FromRgba32(rendered.Width, rendered.Height, rgba),
                OfficeImageExportFormat.Png,
                EncodingOptions);
        }

        return pages;
    }

    private static void RequireInk(byte[] rgba, string label) {
        long ink = 0L;
        for (int index = 0; index < rgba.Length; index += 4) {
            int luma = (rgba[index] * 299 + rgba[index + 1] * 587 + rgba[index + 2] * 114) / 1000;
            if (luma < 224) ink++;
        }

        if (ink * 1000L < rgba.Length / 4) throw new InvalidDataException(label + " rendered blank.");
    }

    // Every benchmark input embeds its fonts.
    private sealed class EmbeddedFontsOnly : IPdfFontResolver {
        public byte[] Resolve(PdfFontRequest request) => null!;
    }
}
#endif
