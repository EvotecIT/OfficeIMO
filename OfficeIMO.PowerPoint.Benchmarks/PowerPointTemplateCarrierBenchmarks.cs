using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointTemplateCarrierBenchmarks {
    private string _directory = string.Empty;
    private string _source = string.Empty;
    private string _cleared = string.Empty;
    private string _retained = string.Empty;
    private string _selected = string.Empty;
    private string _hidden = string.Empty;

    [GlobalSetup]
    public void Setup() {
        _directory = Directory.CreateTempSubdirectory("OfficeIMO-PowerPoint-template-").FullName;
        _source = Path.Combine(_directory, "source.potx");
        _cleared = Path.Combine(_directory, "cleared.pptx");
        _retained = Path.Combine(_directory, "retained.pptx");
        _selected = Path.Combine(_directory, "selected.pptx");
        _hidden = Path.Combine(_directory, "hidden.pptx");
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_TEMPLATE_CARRIER_FIXTURE");
        if (fixture == null) File.WriteAllBytes(_source, CreateFixture());
        else File.Copy(fixture, _source);
        if (CreateWithNoSlides() != 0 || CreateWithSourceSlides() != 24 ||
            CreateWithSelectedHiddenSlides() != 2 || CreateWithAllSlidesHidden() != 24) {
            throw new InvalidOperationException("The converted template did not retain the requested slide inventory.");
        }
        using var document = PowerPointDocument.Load(_retained);
        if (document.SourceFormatDescriptor.Extension != ".pptx" || document.Slides[0].TextBoxes.Count() != 1) {
            throw new InvalidOperationException("The converted template did not preserve the presentation format and authored content.");
        }
        using var selected = PowerPointDocument.Load(_selected);
        using var hidden = PowerPointDocument.Load(_hidden);
        if (selected.Slides.Any(slide => !slide.Hidden) || hidden.Slides.Any(slide => !slide.Hidden)) {
            throw new InvalidOperationException("The converted template did not hide the retained source slides.");
        }
    }

    [GlobalCleanup]
    public void Cleanup() {
        if (Directory.Exists(_directory)) Directory.Delete(_directory, recursive: true);
    }

    [Benchmark]
    public int CreateWithNoSlides() => Create(_cleared, PowerPointTemplateSlideRetention.None);

    [Benchmark]
    public int CreateWithSourceSlides() => Create(_retained, PowerPointTemplateSlideRetention.All);

    [Benchmark]
    public int CreateWithSelectedHiddenSlides() => Create(_selected, PowerPointTemplateSlideRetention.Selected, hide: true);

    [Benchmark]
    public int CreateWithAllSlidesHidden() => Create(_hidden, PowerPointTemplateSlideRetention.All, hide: true);

    private int Create(string output, PowerPointTemplateSlideRetention retention, bool hide = false) {
        var options = new PowerPointTemplateCreationOptions {
            Overwrite = true,
            SlideRetention = retention,
            HideRetainedSourceSlides = hide
        };
        if (retention == PowerPointTemplateSlideRetention.Selected) {
            options.SourceSlideIndexes.Add(0);
            options.SourceSlideIndexes.Add(23);
        }
        using var presentation = PowerPointTemplate.CreateDocument(_source, output, options);
        return presentation.Slides.Count;
    }

    internal static byte[] CreateFixture() => PowerPointModernSubtypeBenchmarks.CreateFixture(PresentationDocumentType.Template);
}
