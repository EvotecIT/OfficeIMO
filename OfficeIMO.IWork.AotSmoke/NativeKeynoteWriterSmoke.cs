using System.Security.Cryptography;
using OfficeIMO.IWork;

internal static class NativeKeynoteWriterSmoke {
    // Exact bytes of the independently decoded, Apple-opened three-slide fixture.
    private const string AcceptedSha256 = "8E25C78B8FFF278DFF256515D6D1A7B171F7C313F82EBCD841457C527F87D6A1";

    internal static void Run() {
        IWorkKeynoteDocument document = CreateAcceptedModel();
        byte[] bytes = document.SaveBytes();
        Require(Convert.ToHexString(SHA256.HashData(bytes)) == AcceptedSha256,
            "NativeAOT writer differs from the independently accepted package.");
        Require(bytes.SequenceEqual(document.SaveBytes()), "NativeAOT writer is nondeterministic.");
        using var stream = new MemoryStream();
        document.Save(stream);
        Require(stream.CanWrite && stream.ToArray().SequenceEqual(bytes), "Stream save differs or closed the caller stream.");

        IWorkKeynoteProjection projection = IWorkSourceDocument.Open(bytes).ReadKeynote();
        Require(projection.HasEditableContent && projection.Slides.Count == 3
            && projection.SlideSize is { WidthPoints: 960, HeightPoints: 540 }, "NativeAOT writer/readback is incomplete.");
        Require(projection.Slides[0].BackgroundColor?.RgbHex == "E6F2FF"
            && projection.Slides[1].BackgroundColor?.RgbHex == "FFE6D9"
            && projection.Slides[2].BackgroundColor?.RgbHex == "FFFFFF", "Written backgrounds differ.");
        Require(projection.Slides[0].TextBoxes.Count == 2 && projection.Slides[1].TextBoxes.Count == 1
            && projection.Slides[2].TextBoxes.Count == 0, "Written text boxes or blank slide differ.");
        Require(projection.Slides[0].TextBoxes[0].Content.PlainText == "OfficeIMO native Keynote\nCreated without a template\n"
            && projection.Slides[0].TextBoxes[1].Content.PlainText == "Times New Roman — 24 pt\n"
            && projection.Slides[1].TextBoxes[0].Content.PlainText == "Zażółć gęślą jaźń\nA😀B — café\n",
            "Written Unicode text or paragraph boundaries differ.");
        Require(projection.Slides[1].TextBoxes[0].Content.Paragraphs[0].Runs[0].Style.Color?.RgbHex == "003366",
            "Written text fill differs.");

        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-Keynote-Aot-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "created.key");
            document.Save(path);
            Require(bytes.SequenceEqual(File.ReadAllBytes(path)), "NativeAOT atomic path save differs.");
            try { document.Save(path); throw new InvalidOperationException("Existing path was replaced."); }
            catch (IOException) { }
            Require(bytes.SequenceEqual(File.ReadAllBytes(path)), "Failed replacement changed the existing path.");
            using var cancelled = new CancellationTokenSource();
            cancelled.Cancel();
            try {
                document.Save(path, overwrite: true, cancellationToken: cancelled.Token);
                throw new InvalidDataException("Pre-cancelled path save completed.");
            } catch (OperationCanceledException) { }
            Require(bytes.SequenceEqual(File.ReadAllBytes(path)), "Cancelled path save changed existing bytes.");
        } finally { Directory.Delete(directory, recursive: true); }

        var larger = IWorkKeynoteDocument.Create();
        larger.AddSlide().AddText(new string('a', 150_000), 0, 0, 960, 540);
        using var activeCancellation = new CancellationTokenSource();
        using var partial = new CancelAfterWriteStream(activeCancellation);
        try {
            larger.Save(partial, cancellationToken: activeCancellation.Token);
            throw new InvalidDataException("Active native-copy cancellation completed.");
        } catch (OperationCanceledException) { }
        Require(partial.CanWrite && partial.Length == 65_536,
            "NativeAOT copy cancellation closed the caller stream or continued beyond its first chunk.");
        Console.WriteLine($"PASS | native Keynote creation | {AcceptedSha256} | {bytes.Length} bytes | deterministic Unicode model, byte/stream/path saves, readback and active cancellation");
    }

    private static IWorkKeynoteDocument CreateAcceptedModel() {
        var document = IWorkKeynoteDocument.Create();
        var first = document.AddSlide("E6F2FF");
        first.AddText("OfficeIMO native Keynote\nCreated without a template", 60, 100, 840, 150, fontSizePoints: 40);
        first.AddText("Times New Roman — 24 pt", 60, 300, 800, 80, "Times New Roman", 24, "003366");
        document.AddSlide("FFE6D9").AddText("Zażółć gęślą jaźń\nA😀B — café", 60, 100, 840, 260, fontSizePoints: 40, color: "003366");
        document.AddSlide();
        return document;
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }

    private sealed class CancelAfterWriteStream(CancellationTokenSource cancellation) : MemoryStream {
        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count);
            cancellation.Cancel();
        }
    }
}
