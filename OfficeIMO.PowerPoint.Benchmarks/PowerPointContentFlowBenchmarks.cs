using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint.Benchmarks;

/// <summary>
/// Flows over richer slide content, once with the Open XML SDK package opened first and once as a loaded deck works by
/// default: edit every chart of a 20-chart deck and save; set every node of a 20-SmartArt deck and save; compose a
/// designed deck from a semantic plan through the layout engine; and export a mixed 20-slide deck (text, tables, charts,
/// SmartArt) to PNG and SVG.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
public class PowerPointContentFlowBenchmarks {
    private const int Slides = 20;
    private byte[] _charts = Array.Empty<byte>();
    private byte[] _smartArt = Array.Empty<byte>();
    private byte[] _mixed = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        _charts = CreateDeck(slide => AddChart(slide));
        _smartArt = CreateDeck(slide => AddSmartArt(slide));
        int index = 0;
        _mixed = CreateDeck(slide => {
            switch (index++ % 4) {
                case 0: AddText(slide); break;
                case 1: AddTable(slide); break;
                case 2: AddChart(slide); break;
                default: AddSmartArt(slide); break;
            }
        });
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("Charts")]
    public byte[] ChartsWithSdk() => Run(_charts, sdk: true, EditCharts);

    [Benchmark]
    [BenchmarkCategory("Charts")]
    public byte[] Charts() => Run(_charts, sdk: false, EditCharts);

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("SmartArt")]
    public byte[] SmartArtWithSdk() => Run(_smartArt, sdk: true, EditSmartArt);

    [Benchmark]
    [BenchmarkCategory("SmartArt")]
    public byte[] SmartArt() => Run(_smartArt, sdk: false, EditSmartArt);

    [Benchmark]
    [BenchmarkCategory("Layout")]
    public byte[] Layout() {
        using PowerPointDocument presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
        presentation.Compose(CreatePlan(), PowerPointCompositionOptions.FromBrief(new PowerPointDesignBrief("#1F4E79", "benchmark")));
        return presentation.ToBytes();
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("ImageExport")]
    public long ImageExportWithSdk() => Export(sdk: true);

    [Benchmark]
    [BenchmarkCategory("ImageExport")]
    public long ImageExport() => Export(sdk: false);

    private static byte[] Run(byte[] deck, bool sdk, Action<PowerPointDocument> flow) {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(deck, writable: false));
        if (sdk) _ = presentation.OpenXmlDocument;
        flow(presentation);
        return presentation.ToBytes();
    }

    private long Export(bool sdk) {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_mixed, writable: false));
        if (sdk) _ = presentation.OpenXmlDocument;
        long bytes = 0;
        foreach (PowerPointSlide slide in presentation.Slides) {
            bytes += slide.ToImage(OfficeImageExportFormat.Png).EncodedLength;
            bytes += slide.ToImage(OfficeImageExportFormat.Svg).EncodedLength;
        }
        return bytes;
    }

    private static void EditCharts(PowerPointDocument presentation) {
        foreach (PowerPointSlide slide in presentation.Slides) {
            foreach (PowerPointChart chart in slide.Shapes.OfType<PowerPointChart>()) {
                chart.SetTitle("Revenue by region");
                chart.SetTitleTextStyle(fontSize: OfficeLength.Points(18), bold: true);
                chart.SetLegend(OfficeChartLegendPosition.Bottom);
                chart.SetValueAxisTitle("EUR thousands");
                chart.SetValueAxisGridlines(showMajor: true);
                chart.SetDataLabels(showValue: true);
                chart.SetSeriesFillColor(0, OfficeColor.FromRgb(0x1F, 0x4E, 0x79));
            }
        }
    }

    private static void EditSmartArt(PowerPointDocument presentation) {
        foreach (PowerPointSlide slide in presentation.Slides) {
            foreach (PowerPointSmartArt smartArt in slide.Shapes.OfType<PowerPointSmartArt>()) {
                for (int node = 0; node < smartArt.NodeCount; node++) smartArt.SetNodeText(node, "Step " + (node + 1));
            }
        }
    }

    private static byte[] CreateDeck(Action<PowerPointSlide> fill) {
        using var stream = new MemoryStream();
        using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
            for (int i = 0; i < Slides; i++) fill(presentation.AddSlide());
            presentation.Save();
        }
        return stream.ToArray();
    }

    private static void AddText(PowerPointSlide slide) =>
        slide.AddTextBox("Quarterly results\nRevenue grew in every region", OfficeLength.Points(40), OfficeLength.Points(40), OfficeLength.Points(600), OfficeLength.Points(120));

    private static void AddTable(PowerPointSlide slide) {
        PowerPointTable table = slide.AddTable(5, 4, OfficeLength.Points(40), OfficeLength.Points(60), OfficeLength.Points(600), OfficeLength.Points(250));
        for (int row = 0; row < 5; row++) {
            for (int column = 0; column < 4; column++) table.GetCell(row, column).Text = row == 0 ? "Region " + column : (row * 100 + column).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
    }

    private static void AddChart(PowerPointSlide slide) =>
        slide.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(
            new[] { "North", "South", "East", "West" },
            new[] {
                new OfficeSeries("2025", new[] { 120d, 95d, 143d, 88d }),
                new OfficeSeries("2026", new[] { 131d, 102d, 150d, 97d })
            }));

    private static void AddSmartArt(PowerPointSlide slide) =>
        slide.AddSmartArt(PowerPointSmartArtType.BasicProcess, new[] { "Discover", "Design", "Pilot", "Roll out" },
            OfficeLength.Points(40), OfficeLength.Points(80), OfficeLength.Points(600), OfficeLength.Points(240));

    private static PowerPointDeckPlan CreatePlan() =>
        new PowerPointDeckPlan()
            .AddSection("Service proposal", "A semantic plan keeps the story reusable while the design can change.", "cover")
            .AddProcess("Implementation path", "The plan describes content intent; the chosen design handles layout.",
                new[] {
                    new PowerPointProcessStep("Discover", "Collect constraints, dependencies, and service expectations."),
                    new PowerPointProcessStep("Design", "Choose target architecture and rollout rules."),
                    new PowerPointProcessStep("Pilot", "Validate the model with a controlled user group."),
                    new PowerPointProcessStep("Roll out", "Deliver in waves with clear reporting.")
                }, "process")
            .AddCardGrid("Scope of services", "Cards are semantic, so the same plan can choose a different grid.",
                new[] {
                    new PowerPointCardContent("Deployments", new[] { "Intune", "Autopilot", "Policy baseline" }),
                    new PowerPointCardContent("Operations", new[] { "Incident handling", "Monitoring", "Optimization" }),
                    new PowerPointCardContent("Consulting", new[] { "Roadmap", "Architecture", "Discovery" }),
                    new PowerPointCardContent("Audits", new[] { "Configuration", "Security review", "Modernization plan" })
                }, "cards")
            .AddCaseStudy("Managed workplace rollout",
                new[] {
                    new PowerPointCaseStudySection("Client", "A distributed organization needed a clear service story."),
                    new PowerPointCaseStudySection("Challenge", "Many locations and mixed hardware made delivery hard to explain."),
                    new PowerPointCaseStudySection("Solution", "Standardized onboarding, monitoring, and operating roles."),
                    new PowerPointCaseStudySection("Result", "Outcomes, metrics, and visual emphasis without hand placement.")
                },
                new[] { new PowerPointMetric("18", "sites"), new PowerPointMetric("420", "devices"), new PowerPointMetric("6", "workstreams") },
                "case-study");

    /// <summary>
    /// Runs each flow <paramref name="rounds"/> times after 10 warm-up rounds, with and without the SDK opened first in
    /// alternating order, and prints median time and allocation. Run it on two revisions in turn for an A/B.
    /// </summary>
    public static int RunInterleaved(int rounds) {
        var benchmark = new PowerPointContentFlowBenchmarks();
        benchmark.Setup();
        // Layout composes a new deck, so it has no SDK-opened variant and reports one result.
        var flows = new (string Name, Func<object>? WithSdk, Func<object> Default)[] {
            ("Charts", benchmark.ChartsWithSdk, benchmark.Charts),
            ("SmartArt", benchmark.SmartArtWithSdk, benchmark.SmartArt),
            ("Layout", null, benchmark.Layout),
            ("ImageExport", () => benchmark.ImageExportWithSdk(), () => benchmark.ImageExport())
        };
        Console.WriteLine("Flow        | SDK opened | default  | SDK alloc | default alloc");
        foreach ((string name, Func<object>? withSdk, Func<object> byDefault) in flows) {
            var times = new[] { new List<double>(), new List<double>() };
            var bytes = new[] { new List<double>(), new List<double>() };
            for (int round = -10; round < rounds; round++) {
                bool sdkFirst = (round & 1) == 0;
                for (int pass = 0; pass < 2; pass++) {
                    int side = (pass == 0) == sdkFirst ? 0 : 1;
                    if (side == 0 && withSdk == null) continue;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    long allocated = GC.GetAllocatedBytesForCurrentThread();
                    long started = System.Diagnostics.Stopwatch.GetTimestamp();
                    _ = side == 0 ? withSdk!() : byDefault();
                    double elapsed = (System.Diagnostics.Stopwatch.GetTimestamp() - started) * 1000D / System.Diagnostics.Stopwatch.Frequency;
                    long used = GC.GetAllocatedBytesForCurrentThread() - allocated;
                    if (round < 0) continue;
                    times[side].Add(elapsed);
                    bytes[side].Add(used);
                }
            }
            Console.WriteLine(withSdk == null
                ? FormattableString.Invariant(
                    $"{name,-11} |          - | {Median(times[1]),6:F2} ms |         - | {Median(bytes[1]) / 1048576D,6:F2} MB")
                : FormattableString.Invariant(
                    $"{name,-11} | {Median(times[0]),7:F2} ms | {Median(times[1]),6:F2} ms | {Median(bytes[0]) / 1048576D,6:F2} MB | {Median(bytes[1]) / 1048576D,6:F2} MB"));
        }
        return 0;
    }

    private static double Median(List<double> values) {
        values.Sort();
        return values.Count == 0 ? 0D : values.Count % 2 == 1 ? values[values.Count / 2] : (values[values.Count / 2 - 1] + values[values.Count / 2]) / 2D;
    }
}
