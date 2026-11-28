using BenchmarkDotNet.Attributes;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointShapeRemovalBenchmarks {
    private const int SlideCount = 20;
    private static readonly Uri Target = new("https://example.test/shape-link");
    private static readonly byte[] Image = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==");
    private byte[] _deck = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_SHAPE_REMOVAL_FIXTURE");
        _deck = string.IsNullOrEmpty(fixture) ? CreateFixture() : File.ReadAllBytes(fixture);
        using var textRemoved = PowerPointDocument.Load(new MemoryStream(RemoveTextShapesAndSave(), writable: false));
        using var imagesRemoved = PowerPointDocument.Load(new MemoryStream(RemovePicturesAndSave(), writable: false));
        if (textRemoved.Slides.Count != SlideCount || textRemoved.Slides.Any(slide =>
                slide.TextBoxes.Single().Hyperlink != Target || slide.Images.Count() != 2)
            || imagesRemoved.Slides.Count != SlideCount || imagesRemoved.Slides.Any(slide =>
                slide.Images.Any() || slide.TextBoxes.Count() != 2)) {
            throw new InvalidOperationException("The benchmark removed the wrong shapes or lost a shared relationship.");
        }
    }

    internal static byte[] CreateFixture() {
        using var presentation = PowerPointDocument.Create();
        for (int index = 0; index < SlideCount; index++) {
            PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.Blank);
            PowerPointTextBox text = slide.AddTextBox("Linked text");
            text.Hyperlink = Target;
            slide.DuplicateShape(text);
            using var image = new MemoryStream(Image, writable: false);
            PowerPointImage picture = slide.AddImage(image, OfficeImageFormat.Png);
            slide.DuplicateShape(picture);
        }
        return presentation.ToBytes();
    }

    [Benchmark]
    public byte[] RemoveTextShapesAndSave() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        foreach (PowerPointSlide slide in presentation.Slides) {
            slide.RemoveShape(slide.TextBoxes.First());
        }
        return presentation.ToBytes();
    }

    [Benchmark]
    public byte[] RemovePicturesAndSave() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        foreach (PowerPointSlide slide in presentation.Slides) {
            foreach (PowerPointImage image in slide.Images.ToArray()) slide.RemoveShape(image);
        }
        return presentation.ToBytes();
    }
}
