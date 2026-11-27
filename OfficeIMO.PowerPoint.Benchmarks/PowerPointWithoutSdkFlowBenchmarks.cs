using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;

namespace OfficeIMO.PowerPoint.Benchmarks;

/// <summary>
/// The realistic flows on the saved 200-slide deck, once with the Open XML SDK package opened, as every load did before,
/// and once without it: open it, add a slide with a text box, an image and a table, and save; open it, edit text, a table
/// cell and the slide order, and save; restyle shapes, build out tables or write speaker notes on every slide, and save;
/// read every layout and add shapes and a picture to the master and a layout, and save;
/// open it, extract its text and close it. Setup checks that the flows without the
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

    /// <summary>Flows whose check setup skips: those that don't stay off the SDK yet in the revision measured.</summary>
    public ISet<string> UncheckedFlows { get; set; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    [GlobalSetup]
    public void Setup() {
        _deck = PowerPointUnchangedSaveWorkload.ReadDeck(Slides);
        _image = CreateImage(640, 480);
        if (!UncheckedFlows.Contains("Add")) Check(Add);
        if (!UncheckedFlows.Contains("Edit")) Check(Edit);
        if (!UncheckedFlows.Contains("TextBoxes")) Check(TextBoxes);
        if (!UncheckedFlows.Contains("Appearance")) Check(Appearance, compareParts: "ppt/slides/slide");
        if (!UncheckedFlows.Contains("Tables")) Check(Tables, compareParts: "ppt/slides/slide");
        if (!UncheckedFlows.Contains("Notes")) Check(Notes, compareParts: "ppt/notesSlides/");
        if (!UncheckedFlows.Contains("Layouts")) Check(Layouts, compareParts: "ppt/slideMasters/slideMaster");
        if (!UncheckedFlows.Contains("Extract")) CheckExtract();
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
    [BenchmarkCategory("Appearance")]
    public byte[] AppearanceWithSdk() => Run(sdk: true, Appearance);

    [Benchmark]
    [BenchmarkCategory("Appearance")]
    public byte[] AppearanceWithoutSdk() => Run(sdk: false, Appearance);

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("Tables")]
    public byte[] TablesWithSdk() => Run(sdk: true, Tables);

    [Benchmark]
    [BenchmarkCategory("Tables")]
    public byte[] TablesWithoutSdk() => Run(sdk: false, Tables);

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("Notes")]
    public byte[] NotesWithSdk() => Run(sdk: true, Notes);

    [Benchmark]
    [BenchmarkCategory("Notes")]
    public byte[] NotesWithoutSdk() => Run(sdk: false, Notes);

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("Layouts")]
    public byte[] LayoutsWithSdk() => Run(sdk: true, Layouts);

    [Benchmark]
    [BenchmarkCategory("Layouts")]
    public byte[] LayoutsWithoutSdk() => Run(sdk: false, Layouts);

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

    // Restyles every shape: a fill with transparency, an outline colour and width, a shadow and soft edges.
    private static void Appearance(PowerPointDocument presentation) {
        foreach (PowerPointSlide slide in presentation.Slides) {
            foreach (PowerPointShape shape in slide.Shapes) {
                if (shape is not (PowerPointAutoShape or PowerPointTextBox)) continue;
                shape.FillColor = OfficeColor.FromRgb(0xDE, 0xEB, 0xF7);
                shape.FillTransparency = 20;
                shape.OutlineColor = OfficeColor.FromRgb(0x1F, 0x4E, 0x79);
                shape.OutlineWidth = OfficeLength.Points(1.5);
                shape.SetShadow(OfficeColor.FromRgb(0, 0, 0), transparencyPercent: 60);
                shape.SetSoftEdges(OfficeLength.Points(1));
            }
        }
    }

    // Adds a table to every slide and builds it out: a bold, filled header row with a border, body text, a merged
    // range, an added row and widths shared out by ratio.
    private static void Tables(PowerPointDocument presentation) {
        foreach (PowerPointSlide slide in presentation.Slides) {
            PowerPointTable table = slide.AddTable(4, 3, OfficeLength.Points(40), OfficeLength.Points(300), OfficeLength.Points(480), OfficeLength.Points(160));
            for (int column = 0; column < 3; column++) {
                PowerPointTableCell header = table.GetCell(0, column);
                header.Text = "Region " + column;
                header.Bold = true;
                header.FillColor = OfficeColor.FromRgb(0x1F, 0x4E, 0x79);
                header.Color = OfficeColor.FromRgb(0xFF, 0xFF, 0xFF);
                header.SetBorders(PowerPointTableCellBorders.Bottom, OfficeColor.FromRgb(0, 0, 0), OfficeLength.Points(1));
            }
            for (int row = 1; row < 4; row++) {
                for (int column = 0; column < 3; column++) table.GetCell(row, column).Text = (row * 100 + column).ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            table.MergeCells(1, 0, 2, 0);
            table.AddRow().Cells[0].Text = "Total";
            table.SetColumnWidthsByRatio(2, 1, 1);
        }
    }

    // Writes two lines of speaker notes on every slide and reads them back.
    private static void Notes(PowerPointDocument presentation) {
        int index = 0;
        foreach (PowerPointSlide slide in presentation.Slides) {
            slide.Notes.Text = "Notes for slide " + index++ + "\nKey point to mention";
            if (!slide.Notes.TryGetText(out _)) throw new InvalidOperationException("The notes were not written.");
        }
    }

    // Lists the layouts and reads every layout's placeholders and every slide's layout, then adds a band and a notice to
    // the master and a logo to a layout.
    private void Layouts(PowerPointDocument presentation) {
        int count = 0;
        foreach (PowerPointSlideLayoutInfo info in presentation.GetSlideLayouts()) count += info.Name.Length;
        foreach (PowerPointSlideMaster master in presentation.SlideMasters) {
            foreach (PowerPointSlideLayout layout in master.Layouts) {
                foreach (PowerPointTextBox placeholder in layout.Placeholders) count += placeholder.PlaceholderType.HasValue ? 1 : 0;
            }
        }
        foreach (PowerPointSlide slide in presentation.Slides) count += slide.GetLayoutPlaceholders().Count + (slide.Layout?.Name.Length ?? 0);
        if (count == 0) throw new InvalidOperationException("The deck has no layouts to read.");
        PowerPointSlideMaster first = presentation.SlideMasters[0];
        first.AddShape(OfficeIMO.Drawing.OfficePresetShapeType.Rectangle, OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(960), OfficeLength.Points(8), "Band");
        first.AddTextBox("Confidential", OfficeLength.Points(20), OfficeLength.Points(510), OfficeLength.Points(200), OfficeLength.Points(20));
        first.Layouts[first.Layouts.Count - 1].AddImage(new MemoryStream(_image, writable: false), OfficeIMO.Drawing.OfficeImageFormat.Png,
            OfficeLength.Points(860), OfficeLength.Points(10), OfficeLength.Points(80), OfficeLength.Points(60));
    }

    // With compareParts the markup of the parts under that prefix is compared, attribute order aside, for flows whose
    // edits the text doesn't show.
    private void Check(Action<PowerPointDocument> flow, string? compareParts = null) {
        string throughSdk, withoutSdk;
        using (PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false))) {
            _ = presentation.OpenXmlDocument;
            flow(presentation);
            byte[] saved = presentation.ToBytes();
            throughSdk = compareParts != null ? PartsXml(saved, compareParts) : Text(saved);
        }
        using (PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false))) {
            flow(presentation);
            byte[] saved = presentation.ToBytes();
            withoutSdk = compareParts != null ? PartsXml(saved, compareParts) : Text(saved);
            if (presentation.IsSdkLoaded) throw new InvalidOperationException("The flow opened the Open XML SDK: " + string.Join("; ", presentation.Diagnostics.Select(diagnostic => diagnostic.Message)));
        }
        if (throughSdk != withoutSdk) throw new InvalidOperationException("The saved slides differ.");
    }

    // Extraction saves nothing, so its check compares the chunks each path extracts.
    private void CheckExtract() {
        string throughSdk, withoutSdk;
        using (PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false))) {
            _ = presentation.OpenXmlDocument;
            throughSdk = string.Join("\n", presentation.ExtractMarkdownChunks().Select(chunk => chunk.Text));
        }
        using (PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false))) {
            withoutSdk = string.Join("\n", presentation.ExtractMarkdownChunks().Select(chunk => chunk.Text));
            if (presentation.IsSdkLoaded) throw new InvalidOperationException("Extraction opened the Open XML SDK: " + string.Join("; ", presentation.Diagnostics.Select(diagnostic => diagnostic.Message)));
        }
        if (throughSdk.Length == 0) throw new InvalidOperationException("Extraction found no text.");
        if (throughSdk != withoutSdk) throw new InvalidOperationException("The extracted chunks differ.");
    }

    /// <summary>
    /// Runs each flow with and without the SDK in alternation, so both see the same machine load, and prints the median
    /// time and allocation of each with their ratios. <paramref name="rounds"/> pairs follow 20 warm-up pairs; each run
    /// starts after a full collection, so neither pays for the other's garbage.
    /// </summary>
    public static int RunInterleaved(int rounds, ISet<string>? uncheckedFlows = null) {
        var benchmark = new PowerPointWithoutSdkFlowBenchmarks { Slides = 200, UncheckedFlows = uncheckedFlows ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase) };
        benchmark.Setup();
        var flows = new (string Name, Func<object> WithSdk, Func<object> WithoutSdk)[] {
            ("Add", benchmark.AddWithSdk, benchmark.AddWithoutSdk),
            ("Edit", benchmark.EditWithSdk, benchmark.EditWithoutSdk),
            ("TextBoxes", benchmark.TextBoxesWithSdk, benchmark.TextBoxesWithoutSdk),
            ("Appearance", benchmark.AppearanceWithSdk, benchmark.AppearanceWithoutSdk),
            ("Tables", benchmark.TablesWithSdk, benchmark.TablesWithoutSdk),
            ("Notes", benchmark.NotesWithSdk, benchmark.NotesWithoutSdk),
            ("Layouts", benchmark.LayoutsWithSdk, benchmark.LayoutsWithoutSdk),
            ("Extract", () => benchmark.ExtractWithSdk(), () => benchmark.ExtractWithoutSdk())
        };
        Console.WriteLine("Flow       | SDK median | without | faster | SDK alloc | without | less allocation");
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
            Console.WriteLine(FormattableString.Invariant($"{name,-10} | {sdkTime,7:F2} ms | {ownTime,5:F2} ms | {sdkTime / ownTime,5:F2}x | {sdkAlloc / 1048576D,6:F2} MB | {ownAlloc / 1048576D,5:F2} MB | {sdkAlloc / ownAlloc,5:F2}x"));
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

    private static string PartsXml(byte[] package, string prefix) {
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(package, writable: false));
        var builder = new System.Text.StringBuilder();
        foreach (System.IO.Compression.ZipArchiveEntry entry in archive.Entries
                     .Where(entry => entry.FullName.StartsWith(prefix, StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
                     .OrderBy(entry => entry.FullName, StringComparer.Ordinal)) {
            using Stream stream = entry.Open();
            System.Xml.Linq.XElement root = System.Xml.Linq.XElement.Load(stream);
            foreach (System.Xml.Linq.XElement element in root.DescendantsAndSelf()) {
                System.Xml.Linq.XAttribute[] attributes = element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
                    .OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal).ToArray();
                element.RemoveAttributes();
                element.Add(attributes);
            }
            builder.Append(entry.FullName).Append('\n').Append(root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)).Append('\n');
        }
        return builder.ToString();
    }

    private static string Text(byte[] package) {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(package, writable: false));
        return string.Concat(presentation.ExtractMarkdownChunks().Select(chunk => chunk.Text))
            + string.Join("|", presentation.Slides.Select(slide => slide.Shapes.Count));
    }
}
