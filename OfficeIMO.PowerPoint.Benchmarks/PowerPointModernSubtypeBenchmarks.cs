using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointModernSubtypeBenchmarks {
    private const int SlideCount = 24;
    private const string SlideText = "PowerPoint subtype fixture";
    private byte[] _template = Array.Empty<byte>();
    private byte[] _slideshow = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? directory = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_SUBTYPE_FIXTURES");
        _template = directory == null
            ? CreateFixture(PresentationDocumentType.Template)
            : File.ReadAllBytes(Path.Combine(directory, "fixture.potx"));
        _slideshow = directory == null
            ? CreateFixture(PresentationDocumentType.MacroEnabledSlideshow)
            : File.ReadAllBytes(Path.Combine(directory, "fixture.ppsm"));
        if (ReadTemplateText() != SlideCount * SlideText.Length || ReadSlideshowText() != SlideCount * SlideText.Length) {
            throw new InvalidOperationException("The benchmark fixture lost its slides or authored text.");
        }
    }

    [Benchmark]
    public int ReadTemplateText() => ReadSlideText(_template, ".potx");

    [Benchmark]
    public int ReadSlideshowText() => ReadSlideText(_slideshow, ".ppsm");

    private static int ReadSlideText(byte[] bytes, string extension) {
        using var source = new MemoryStream(bytes, writable: false);
        using var presentation = PowerPointDocument.Load(source, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        if (presentation.SourceFormatDescriptor.Extension != extension)
            throw new InvalidOperationException("The benchmark did not open the expected presentation subtype.");
        return presentation.Slides.Sum(slide => slide.TextBoxes.Single().Text.Length);
    }

    internal static byte[] CreateFixture(PresentationDocumentType type) {
        using var presentation = PowerPointDocument.Create();
        for (int index = 0; index < SlideCount; index++) {
            presentation.AddSlide(PowerPointSlideLayoutType.Blank).AddTextBox(SlideText);
        }
        return ChangeSubtype(presentation.ToBytes(), type);
    }

    internal static byte[] ChangeSubtype(byte[] bytes, PresentationDocumentType type) {
        using var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        using (var package = PresentationDocument.Open(stream, true)) {
            package.ChangeDocumentType(type);
        }
        return stream.ToArray();
    }
}
