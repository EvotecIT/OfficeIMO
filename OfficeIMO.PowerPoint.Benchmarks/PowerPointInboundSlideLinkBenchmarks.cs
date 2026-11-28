using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointInboundSlideLinkBenchmarks {
    private const int RemovedIndex = 7;
    private byte[] _shapeLinks = Array.Empty<byte>();
    private byte[] _textLinks = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? directory = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_INBOUND_LINK_FIXTURES");
        _shapeLinks = directory == null ? CreateFixture(textLinks: false) : File.ReadAllBytes(Path.Combine(directory, "shapes.pptx"));
        _textLinks = directory == null ? CreateFixture(textLinks: true) : File.ReadAllBytes(Path.Combine(directory, "runs.pptx"));
        Verify(RemoveSlideAndShapeLinks(), textLinks: false);
        Verify(RemoveSlideAndTextLinks(), textLinks: true);
    }

    private static void Verify(byte[] bytes, bool textLinks) {
        using var document = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        if (document.Slides.Count != 19) throw new InvalidOperationException("The saved deck did not remove the requested slide.");
        for (int index = 0; index < 3; index++) {
            PowerPointSlide slide = document.Slides[index];
            Uri? removedLink;
            Uri? retainedLink;
            if (textLinks) {
                PowerPointTextBox[] boxes = slide.TextBoxes.ToArray();
                removedLink = boxes[1].Paragraphs.Single().Runs.Single().Hyperlink;
                retainedLink = boxes[2].Paragraphs.Single().Runs.Single().Hyperlink;
            } else {
                removedLink = slide.Shapes[1].Hyperlink;
                retainedLink = slide.Shapes[2].Hyperlink;
            }

            if (removedLink != null || retainedLink?.OriginalString != "#slide-10") {
                throw new InvalidOperationException("The saved deck did not preserve the surviving internal hyperlink.");
            }
        }
    }

    [Benchmark]
    public byte[] RemoveSlideAndShapeLinks() => RemoveAndSave(_shapeLinks);

    [Benchmark]
    public byte[] RemoveSlideAndTextLinks() => RemoveAndSave(_textLinks);

    private static byte[] RemoveAndSave(byte[] bytes) {
        using var document = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        document.RemoveSlide(RemovedIndex);
        return document.ToBytes();
    }

    internal static byte[] CreateFixture(bool textLinks) {
        using var document = PowerPointDocument.Create();
        for (int index = 0; index < 20; index++) {
            document.AddSlide(PowerPointSlideLayoutType.Blank).AddTextBox("Planning slide " + index);
        }
        AddFixtureLinks(document, RemovedIndex, 10, textLinks);
        return document.ToBytes();
    }

    internal static void AddFixtureLinks(PowerPointDocument document, int removedIndex, int retainedIndex, bool textLinks) {
        PowerPointSlide removed = document.Slides[removedIndex];
        PowerPointSlide retained = document.Slides[retainedIndex];
        foreach (PowerPointSlide source in document.Slides.Where(slide => slide != removed && slide != retained).Take(3)) {
            if (textLinks) {
                source.AddTextBox("Removed target").Paragraphs.Single().Runs.Single().SetHyperlink(removed);
                source.AddTextBox("Retained target").Paragraphs.Single().Runs.Single().SetHyperlink(retained);
            } else {
                source.AddRectangle(OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(120), OfficeLength.Points(44)).SetHyperlink(removed);
                source.AddRectangle(OfficeLength.Points(20), OfficeLength.Points(80), OfficeLength.Points(120), OfficeLength.Points(44)).SetHyperlink(retained);
            }
        }
    }
}
