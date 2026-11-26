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

    /// <summary>Whether setup checks that the text box flow stays off the SDK; off for revisions where it doesn't yet.</summary>
    public bool CheckTextBoxes { get; set; } = true;

    [GlobalSetup]
    public void Setup() {
        _deck = PowerPointUnchangedSaveWorkload.ReadDeck(Slides);
        _image = CreateImage(640, 480);
        Check(Add);
        Check(Edit);
        if (CheckTextBoxes) Check(TextBoxes);
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
    [BenchmarkCategory("TextBoxes")]
    public byte[] TextBoxesWithSdk() => Run(sdk: true, TextBoxes);

    [Benchmark]
    [BenchmarkCategory("TextBoxes")]
    public byte[] TextBoxesWithoutSdk() => Run(sdk: false, TextBoxes);

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

    // Rewrites and formats the text boxes of every slide: two paragraphs, run formatting, a colour, margins and auto-fit.
    private static void TextBoxes(PowerPointDocument presentation) {
        foreach (PowerPointSlide slide in presentation.Slides) {
            foreach (PowerPointTextBox box in slide.TextBoxes) {
                box.Text = "Summary\nDetails follow";
                box.Bold = true;
                box.FontSize = OfficeLength.Points(20);
                box.Color = OfficeColor.FromRgb(0x1F, 0x4E, 0x79);
                box.SetTextMargins(OfficeLength.Points(6), OfficeLength.Points(3), OfficeLength.Points(6), OfficeLength.Points(3));
                box.TextAutoFit = PowerPointTextAutoFit.Shape;
                box.Paragraphs[1].AddText(" and more").Italic = true;
            }
        }
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

    /// <summary>
    /// Runs each flow with and without the SDK in alternation, so both see the same machine load, and prints the median
    /// time and allocation of each with their ratios. <paramref name="rounds"/> pairs follow 20 warm-up pairs; each run
    /// starts after a full collection, so neither pays for the other's garbage.
    /// </summary>
    public static int RunInterleaved(int rounds, bool checkTextBoxes = true) {
        var benchmark = new PowerPointWithoutSdkFlowBenchmarks { Slides = 200, CheckTextBoxes = checkTextBoxes };
        benchmark.Setup();
        var flows = new (string Name, Func<object> WithSdk, Func<object> WithoutSdk)[] {
            ("Add", benchmark.AddWithSdk, benchmark.AddWithoutSdk),
            ("Edit", benchmark.EditWithSdk, benchmark.EditWithoutSdk),
            ("TextBoxes", benchmark.TextBoxesWithSdk, benchmark.TextBoxesWithoutSdk),
            ("Extract", () => benchmark.ExtractWithSdk(), () => benchmark.ExtractWithoutSdk())
        };
        Console.WriteLine("Flow      | SDK median | without | faster | SDK alloc | without | less allocation");
        foreach ((string name, Func<object> withSdk, Func<object> withoutSdk) in flows) {
            var sdkTimes = new List<double>(); var ownTimes = new List<double>();
            var sdkBytes = new List<long>(); var ownBytes = new List<long>();
            for (int round = -20; round < rounds; round++) {
                // Alternate which side goes first, so neither always runs on a warmer cache.
                bool sdkFirst = (round & 1) == 0;
                for (int side = 0; side < 2; side++) {
                    bool sdk = side == 0 ? sdkFirst : !sdkFirst;
                    // Collections the other side's garbage would trigger are paid here, outside the timing.
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    long allocated = GC.GetAllocatedBytesForCurrentThread();
                    long started = System.Diagnostics.Stopwatch.GetTimestamp();
                    _ = sdk ? withSdk() : withoutSdk();
                    double elapsed = (System.Diagnostics.Stopwatch.GetTimestamp() - started) * 1000D / System.Diagnostics.Stopwatch.Frequency;
                    long bytes = GC.GetAllocatedBytesForCurrentThread() - allocated;
                    if (round < 0) continue;
                    (sdk ? sdkTimes : ownTimes).Add(elapsed);
                    (sdk ? sdkBytes : ownBytes).Add(bytes);
                }
            }
            double sdkTime = Median(sdkTimes), ownTime = Median(ownTimes);
            double sdkAlloc = Median(sdkBytes.Select(value => (double)value).ToList()), ownAlloc = Median(ownBytes.Select(value => (double)value).ToList());
            Console.WriteLine(FormattableString.Invariant($"{name,-9} | {sdkTime,7:F2} ms | {ownTime,5:F2} ms | {sdkTime / ownTime,5:F2}x | {sdkAlloc / 1048576D,6:F2} MB | {ownAlloc / 1048576D,5:F2} MB | {sdkAlloc / ownAlloc,5:F2}x"));
        }
        return 0;
    }

    private static double Median(List<double> values) {
        values.Sort();
        return values.Count == 0 ? 0D : values.Count % 2 == 1 ? values[values.Count / 2] : (values[values.Count / 2 - 1] + values[values.Count / 2]) / 2D;
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
