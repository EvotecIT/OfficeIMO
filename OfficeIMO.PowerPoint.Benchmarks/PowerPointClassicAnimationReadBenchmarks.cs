using System.IO.Compression;
using System.Xml.Linq;
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointClassicAnimationReadBenchmarks {
    private const int SlideCount = 24;
    private const int Delay = 400;
    private static readonly XNamespace Presentation = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace Producer = "https://schemas.officeimo.net/powerpoint/2026/classic-animations";
    private byte[] _deck = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        string? fixture = Environment.GetEnvironmentVariable("OFFICEIMO_POWERPOINT_CLASSIC_READ_FIXTURE");
        _deck = string.IsNullOrEmpty(fixture) ? CreateFixture() : File.ReadAllBytes(fixture);
        if (ReadAuthoredAnimations() != SlideCount / 2 * Delay || ReadInferredAnimations() != SlideCount / 2 * Delay) {
            throw new InvalidOperationException("The classic timing benchmark fixture has unexpected animation data.");
        }
    }

    internal static byte[] CreateFixture() {
        using var deck = PowerPointDocument.Create();
        for (int index = 0; index < SlideCount; index++) {
            PowerPointSlide slide = deck.AddSlide(PowerPointSlideLayoutType.Blank);
            slide.AddClassicAnimation(slide.AddTextBox("Classic entrance effect"), PowerPointClassicAnimationEffect.Fade,
                new PowerPointClassicAnimationOptions {
                    Automatic = true, DelayMilliseconds = Delay, BuildType = PowerPointClassicAnimationBuildType.ByLevel2Paragraph
                });
        }
        using var bytes = new MemoryStream();
        deck.Save(bytes);
        bytes.Position = 0;
        using (var archive = new ZipArchive(bytes, ZipArchiveMode.Update, leaveOpen: true)) {
            for (int slideNumber = 2; slideNumber <= SlideCount; slideNumber += 2) {
                string part = $"ppt/slides/slide{slideNumber}.xml";
                ZipArchiveEntry entry = archive.GetEntry(part)!;
                XDocument xml;
                using (Stream input = entry.Open()) xml = XDocument.Load(input);
                xml.Root!.Element(Presentation + "extLst")!.Elements(Presentation + "ext")
                    .Single(extension => extension.Element(Producer + "classicAnimations") != null).Remove();
                entry.Delete();
                using Stream output = archive.CreateEntry(part).Open();
                xml.Save(output);
            }
        }
        return bytes.ToArray();
    }

    [Benchmark]
    public int ReadAuthoredAnimations() => ReadAnimations(0);

    [Benchmark]
    public int ReadInferredAnimations() => ReadAnimations(1);

    private int ReadAnimations(int firstSlide) {
        using var presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        int delays = 0;
        for (int index = firstSlide; index < presentation.Slides.Count; index += 2) {
            delays += presentation.Slides[index].ClassicAnimations.Single().DelayMilliseconds;
        }
        return delays;
    }
}
