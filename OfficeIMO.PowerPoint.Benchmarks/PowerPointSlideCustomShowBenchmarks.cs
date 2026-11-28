using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointSlideCustomShowBenchmarks {
    private byte[] _shared = Array.Empty<byte>();
    private byte[] _linked = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? directory = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_SLIDE_CUSTOM_SHOW_FIXTURES");
        _shared = directory == null ? CreateFixture(linked: false) : File.ReadAllBytes(Path.Combine(directory, "shared.pptx"));
        _linked = directory == null ? CreateFixture(linked: true) : File.ReadAllBytes(Path.Combine(directory, "linked.pptx"));
        Verify(RemoveSlideFromSharedShows(), linked: false);
        Verify(RemoveSlideAndLinkedShow(), linked: true);
    }

    private static void Verify(byte[] bytes, bool linked) {
        using var document = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        PowerPointCustomShow mixed = document.GetCustomShow("Mixed")!;
        if (document.Slides.Count != 19 || document.CustomShows.Count != (linked ? 1 : 2) ||
            !mixed.Slides.SequenceEqual(new[] { document.Slides[0], document.Slides[8] })) {
            throw new InvalidOperationException("The saved fixture did not prune the custom-show slide sequences.");
        }
        SlidePart first = document.Slides[0].SlidePart;
        if (first.HyperlinkRelationships.Any(edge => edge.Id == "rIdPrunedShow") ||
            first.Slide!.Descendants<A.HyperlinkType>().Any(link => link.Action?.Value == "ppaction://customshow?id=2")) {
            throw new InvalidOperationException("The saved fixture retained a link to the emptied show.");
        }
    }

    [Benchmark]
    public byte[] RemoveSlideFromSharedShows() => RemoveAndSave(_shared);

    [Benchmark]
    public byte[] RemoveSlideAndLinkedShow() => RemoveAndSave(_linked);

    private static byte[] RemoveAndSave(byte[] bytes) {
        using var document = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        document.RemoveSlide(7);
        return document.ToBytes();
    }

    internal static byte[] CreateFixture(bool linked) {
        using var document = PowerPointDocument.Create();
        for (int index = 0; index < 20; index++) {
            document.AddSlide(PowerPointSlideLayoutType.Blank).AddTextBox("Planning slide " + index);
        }
        AddFixtureShows(document, 7, "Mixed", "Dated", linked);
        return document.ToBytes();
    }

    internal static void AddFixtureShows(PowerPointDocument document, int removedIndex, string mixedName, string datedName, bool linked) {
        IReadOnlyList<PowerPointSlide> slides = document.Slides;
        PowerPointSlide removed = slides[removedIndex];
        PowerPointSlide last = slides[Math.Min(slides.Count - 1, removedIndex + 2)];
        document.AddCustomShow(mixedName, new[] { removed, slides[0], removed, last });
        PowerPointCustomShow dated = document.AddCustomShow(datedName,
            linked ? new[] { removed, removed } : new[] { removed, last, removed });
        if (!linked) return;
        SlidePart part = slides[0].SlidePart;
        part.AddHyperlinkRelationship(new Uri("https://example.test/pruned-show"), true, "rIdPrunedShow");
        PowerPointAutoShape rectangle = slides[0].AddRectangle(OfficeLength.Points(10), OfficeLength.Points(20),
            OfficeLength.Points(100), OfficeLength.Points(40));
        var shape = (P.Shape)rectangle.Element;
        shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Append(new A.HyperlinkOnClick {
            Id = "rIdPrunedShow",
            Action = "ppaction://customshow?id=" + dated.Id
        });
    }
}
