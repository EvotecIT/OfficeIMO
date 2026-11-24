using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Core.Internal.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Reader.Benchmarks;

/// <summary>
/// Reads the text of every DrawingML paragraph on every slide, in presentation order, from the deck that
/// <see cref="PowerPointSlideStreamingBenchmarks"/> uses: directly through the Open XML SDK, and through the
/// SDK-free package and pull reader in OfficeIMO.Core. Setup checks that both lanes return the same paragraphs.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointSlideTextBenchmarks {
    private byte[] _package = Array.Empty<byte>();

    [Params(20, 200)]
    public int Slides { get; set; }

    [GlobalSetup]
    public void Setup() {
        string path = PowerPointStreamingDeck.Create(Slides);
        try {
            _package = File.ReadAllBytes(path);
        } finally {
            PowerPointStreamingDeck.Delete(path);
        }
        TextObservation sdk = OpenXmlSdk();
        TextObservation packageReader = PackageReader();
        if (sdk.Paragraphs == 0 || sdk != packageReader) {
            throw new InvalidOperationException("The package reader observed " + packageReader + "; the SDK observed " + sdk + ".");
        }
    }

    [Benchmark(Baseline = true)]
    public TextObservation OpenXmlSdk() {
        using var stream = new MemoryStream(_package, writable: false);
        using PresentationDocument document = PresentationDocument.Open(stream, false);
        PresentationPart presentation = document.PresentationPart!;
        var observation = TextObservation.Empty;
        int index = 0;
        foreach (SlideId slideId in presentation.Presentation!.SlideIdList!.Elements<SlideId>()) {
            var slide = (SlidePart)presentation.GetPartById(slideId.RelationshipId!.Value!);
            foreach (A.Paragraph paragraph in slide.Slide!.Descendants<A.Paragraph>()) {
                observation = observation.Add(index, paragraph.InnerText);
            }
            index++;
        }
        return observation;
    }

    [Benchmark]
    public TextObservation PackageReader() {
        using var stream = new MemoryStream(_package, writable: false);
        var observation = TextObservation.Empty;
        foreach (OfficeSlideParagraph paragraph in OfficeOpenXmlTextReader.ReadSlideParagraphs(stream)) {
            observation = observation.Add(paragraph.SlideIndex, paragraph.Text);
        }
        return observation;
    }

    /// <summary>Paragraph and character counts with an order-sensitive checksum of slide positions and text.</summary>
    public readonly record struct TextObservation(int Paragraphs, int Characters, ulong Checksum) {
        public static TextObservation Empty => new(0, 0, 14695981039346656037UL);

        public TextObservation Add(int slide, string text) {
            ulong checksum = (Checksum ^ (ulong)slide) * 1099511628211UL;
            foreach (char character in text) {
                checksum = (checksum ^ character) * 1099511628211UL;
            }
            return new TextObservation(Paragraphs + 1, Characters + text.Length, (checksum ^ 0xFF) * 1099511628211UL);
        }
    }
}
