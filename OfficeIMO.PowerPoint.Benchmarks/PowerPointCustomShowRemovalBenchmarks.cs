using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointCustomShowRemovalBenchmarks {
    private byte[] _unlinked = Array.Empty<byte>();
    private byte[] _linked = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? directory = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_CUSTOM_SHOW_REMOVAL_FIXTURES");
        _unlinked = directory == null ? CreateFixture(linked: false) : File.ReadAllBytes(Path.Combine(directory, "unlinked.pptx"));
        _linked = directory == null ? CreateFixture(linked: true) : File.ReadAllBytes(Path.Combine(directory, "linked.pptx"));
        Verify(RemoveUnlinkedShowAndSave());
        Verify(RemoveLinkedShowAndSave());
    }

    private static void Verify(byte[] bytes) {
        using var document = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        if (document.CustomShows.Count != 1 || document.CustomShows[0].Name != "Keep" || document.Slides.Count != 20) {
            throw new InvalidOperationException("The saved fixture did not preserve the remaining show and slides.");
        }
        SlidePart first = document.Slides[0].SlidePart;
        if (first.HyperlinkRelationships.Any(edge => edge.Id == "rIdRemovalExternal") ||
            first.Parts.Any(edge => edge.RelationshipId == "rIdRemovalInternal") ||
            first.Slide!.Descendants<A.HyperlinkType>().Any(link => link.Action?.Value == "ppaction://customshow?id=1")) {
            throw new InvalidOperationException("The saved fixture retained an action linked to the removed show.");
        }
    }

    [Benchmark]
    public byte[] RemoveUnlinkedShowAndSave() => RemoveAndSave(_unlinked);

    [Benchmark]
    public byte[] RemoveLinkedShowAndSave() => RemoveAndSave(_linked);

    private static byte[] RemoveAndSave(byte[] bytes) {
        using var document = PowerPointDocument.Load(new MemoryStream(bytes, writable: false));
        document.RemoveCustomShow(document.GetCustomShow("Remove")!);
        return document.ToBytes();
    }

    internal static byte[] CreateFixture(bool linked) {
        using var document = PowerPointDocument.Create();
        for (int index = 0; index < 20; index++) {
            document.AddSlide(PowerPointSlideLayoutType.Blank).AddTextBox("Decision slide " + index);
        }
        PowerPointCustomShow show = document.AddCustomShow("Remove", document.Slides.Take(6));
        document.AddCustomShow("Keep", document.Slides.Skip(6).Take(8));
        if (linked) AddFixtureActions(document, show);
        return document.ToBytes();
    }

    internal static void AddFixtureActions(PowerPointDocument document, PowerPointCustomShow show) {
        PowerPointSlide slide = document.Slides[0];
        slide.SlidePart.AddHyperlinkRelationship(new Uri("https://example.test/custom-show-removal"), true, "rIdRemovalExternal");
        string internalId = slide.SlidePart.Parts
            .FirstOrDefault(pair => ReferenceEquals(pair.OpenXmlPart, document.Slides[1].SlidePart)).RelationshipId ?? "rIdRemovalInternal";
        if (internalId == "rIdRemovalInternal") slide.SlidePart.AddPart(document.Slides[1].SlidePart, internalId);
        int index = 0;
        foreach (string id in new[] { "rIdRemovalExternal", internalId }) {
            PowerPointAutoShape rectangle = slide.AddRectangle(OfficeLength.Points(10), OfficeLength.Points(20 + 50 * index++),
                OfficeLength.Points(100), OfficeLength.Points(40));
            var shape = (P.Shape)rectangle.Element;
            shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Append(new A.HyperlinkOnClick {
                Id = id,
                Action = "ppaction://customshow?id=" + show.Id
            });
        }
    }
}
