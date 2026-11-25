using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;

namespace OfficeIMO.PowerPoint.Benchmarks;

/// <summary>
/// The realistic flows on the saved 200-slide deck, once with the Open XML SDK package opened, as every load did before,
/// and once without it: open it, add a slide with a text box, an image and a table, and save; open it, edit text, a table
/// cell and the slide order, and save; open it, extract its text and close it. Setup checks that the flows without the
/// SDK never open it and save the same slides as through it.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
public class PowerPointWithoutSdkFlowBenchmarks {
    private byte[] _deck = Array.Empty<byte>();
    private byte[] _image = Array.Empty<byte>();

    [Params(200)]
    public int Slides { get; set; }

    [GlobalSetup]
    public void Setup() {
        _deck = PowerPointUnchangedSaveWorkload.ReadDeck(Slides);
        _image = CreateImage(640, 480);
        Check(Add);
        Check(Edit);
        Check(presentation => presentation.ExtractMarkdownChunks().Count());
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("Add")]
    public byte[] AddWithSdk() => Run(sdk: true, Add);

    [Benchmark]
    [BenchmarkCategory("Add")]
    public byte[] AddWithoutSdk() => Run(sdk: false, Add);

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("Edit")]
    public byte[] EditWithSdk() => Run(sdk: true, Edit);

    [Benchmark]
    [BenchmarkCategory("Edit")]
    public byte[] EditWithoutSdk() => Run(sdk: false, Edit);

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("Extract")]
    public int ExtractWithSdk() => Extract(sdk: true);

    [Benchmark]
    [BenchmarkCategory("Extract")]
    public int ExtractWithoutSdk() => Extract(sdk: false);

    private byte[] Run(bool sdk, Action<PowerPointDocument> flow) {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        if (sdk) _ = presentation.OpenXmlDocument;
        flow(presentation);
        return presentation.ToBytes();
    }

    private int Extract(bool sdk) {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        if (sdk) _ = presentation.OpenXmlDocument;
        return presentation.ExtractMarkdownChunks().Count();
    }

    private void Add(PowerPointDocument presentation) {
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTextBox("Quarterly results", OfficeLength.Points(40), OfficeLength.Points(30), OfficeLength.Points(600), OfficeLength.Points(50));
        slide.AddImage(new MemoryStream(_image, writable: false), OfficeIMO.Drawing.OfficeImageFormat.Png,
            OfficeLength.Points(40), OfficeLength.Points(100), OfficeLength.Points(200), OfficeLength.Points(150));
        PowerPointTable table = slide.AddTable(4, 3, OfficeLength.Points(260), OfficeLength.Points(100), OfficeLength.Points(400), OfficeLength.Points(160));
        for (int row = 0; row < 4; row++) {
            for (int column = 0; column < 3; column++) table.GetCell(row, column).Text = row == 0 ? "Region " + column : (row * 100 + column).ToString();
        }
    }

    private static void Edit(PowerPointDocument presentation) {
        PowerPointTextBox box = presentation.Slides[3].TextBoxes.First();
        box.Text = "Edited title";
        presentation.ReplaceAll("point", "item");
        presentation.RemoveSlide(10);
        presentation.MoveSlide(20, 0);
    }

    private void Check(Action<PowerPointDocument> flow) {
        string throughSdk, withoutSdk;
        using (PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false))) {
            _ = presentation.OpenXmlDocument;
            flow(presentation);
            throughSdk = Text(presentation.ToBytes());
        }
        using (PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false))) {
            flow(presentation);
            withoutSdk = Text(presentation.ToBytes());
            if (presentation.IsSdkLoaded) throw new InvalidOperationException("The flow opened the Open XML SDK: " + string.Join("; ", presentation.Diagnostics.Select(diagnostic => diagnostic.Message)));
        }
        if (throughSdk != withoutSdk) throw new InvalidOperationException("The saved slides differ.");
    }

    // A gradient photo-sized PNG, so the image part has realistic weight.
    private static byte[] CreateImage(int width, int height) {
        byte[] rgba = new byte[width * height * 4];
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                int offset = (y * width + x) * 4;
                rgba[offset] = (byte)(x * 255 / width);
                rgba[offset + 1] = (byte)(y * 255 / height);
                rgba[offset + 2] = 128;
                rgba[offset + 3] = 255;
            }
        }
        return OfficeIMO.Drawing.OfficePngWriter.EncodeRgba(width, height, rgba);
    }

    private static string Text(byte[] package) {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(package, writable: false));
        return string.Concat(presentation.ExtractMarkdownChunks().Select(chunk => chunk.Text))
            + string.Join("|", presentation.Slides.Select(slide => slide.Shapes.Count));
    }
}
