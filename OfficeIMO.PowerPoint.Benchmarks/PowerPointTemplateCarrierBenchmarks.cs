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

    [GlobalSetup]
    public void Setup() {
        _directory = Directory.CreateTempSubdirectory("OfficeIMO-PowerPoint-template-").FullName;
        _source = Path.Combine(_directory, "source.potx");
        _cleared = Path.Combine(_directory, "cleared.pptx");
        _retained = Path.Combine(_directory, "retained.pptx");
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_TEMPLATE_CARRIER_FIXTURE");
        if (fixture == null) File.WriteAllBytes(_source, CreateFixture());
        else File.Copy(fixture, _source);
        if (CreateWithNoSlides() != 0 || CreateWithSourceSlides() != 24) {
            throw new InvalidOperationException("The converted template did not retain the requested slide inventory.");
        }
        using var document = PowerPointDocument.Load(_retained);
        if (document.SourceFormatDescriptor.Extension != ".pptx" || document.Slides[0].TextBoxes.Count() != 1) {
            throw new InvalidOperationException("The converted template did not preserve the presentation format and authored content.");
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

    private int Create(string output, PowerPointTemplateSlideRetention retention) {
        using var presentation = PowerPointTemplate.CreateDocument(_source, output, new PowerPointTemplateCreationOptions {
            Overwrite = true,
            SlideRetention = retention
        });
        return presentation.Slides.Count;
    }

    internal static byte[] CreateFixture() => PowerPointModernSubtypeBenchmarks.CreateFixture(PresentationDocumentType.Template);
}
