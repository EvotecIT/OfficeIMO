using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>
/// Loads a DOCX of distinct images from memory, read-only and editable, at about 1 MB and 10 MB.
/// Allocation divided by <see cref="InputBytes"/> shows how many copies of the package a load makes.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "Load")]
public class WordLoadBenchmarks {
    private byte[] _document = Array.Empty<byte>();

    [Params(50, 500)]
    public int ImageCount { get; set; }

    [Params(DocumentAccessMode.ReadOnly, DocumentAccessMode.ReadWrite)]
    public DocumentAccessMode AccessMode { get; set; }

    internal int InputBytes => _document.Length;

    [GlobalSetup]
    public void Setup() {
        using (WordDocument document = WordDocument.Create()) {
            for (int index = 0; index < ImageCount; index++) {
                document.AddParagraph(WordWorkloadCorpus.ImageCaption(index));
                using var image = new MemoryStream(WordBenchmarkImages.Image(index), writable: false);
                document.AddParagraph().AddImage(image, "image" + index + ".png",
                    OfficeLength.Pixels(WordBenchmarkImages.WidthPixels),
                    OfficeLength.Pixels(WordBenchmarkImages.HeightPixels));
            }
            _document = document.ToBytes();
        }

        int images = Load();
        if (images != ImageCount) {
            throw new InvalidOperationException("Loading found " + images + " images; expected " + ImageCount + ".");
        }
        Console.WriteLine("// Input bytes: " + InputBytes);
    }

    [Benchmark]
    public int Load() {
        using var stream = new MemoryStream(_document, writable: false);
        using WordDocument document = WordDocument.Load(stream, new WordLoadOptions { AccessMode = AccessMode });
        return document.Images.Count;
    }
}
