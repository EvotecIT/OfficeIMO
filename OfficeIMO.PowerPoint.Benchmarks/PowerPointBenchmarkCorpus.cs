using OfficeIMO.PowerPoint;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint.Benchmarks;

internal sealed record PowerPointBenchmarkFixture(
    string Scale,
    int SlideCount,
    int ExpectedMinimumShapeCount);

internal static class PowerPointBenchmarkCorpus {
    internal static readonly IReadOnlyList<string> Scales =
        new[] { "Small", "Normal", "Large" };

    internal static PowerPointBenchmarkFixture Get(string scale) {
        if (string.Equals(scale, "Small", StringComparison.OrdinalIgnoreCase)) {
            return new PowerPointBenchmarkFixture("Small", 3, 18);
        }
        if (string.Equals(scale, "Normal", StringComparison.OrdinalIgnoreCase)) {
            return new PowerPointBenchmarkFixture("Normal", 30, 180);
        }
        if (string.Equals(scale, "Large", StringComparison.OrdinalIgnoreCase)) {
            return new PowerPointBenchmarkFixture("Large", 120, 720);
        }
        throw new ArgumentException(
            "Scale must be Small, Normal, or Large.", nameof(scale));
    }

    internal static byte[] CreatePackage(PowerPointBenchmarkFixture fixture) {
        using var stream = new MemoryStream();
        using (PowerPointDocument presentation =
               PowerPointDocument.Create(stream)) {
            Populate(presentation, fixture);
            presentation.Save();
        }
        return stream.ToArray();
    }

    internal static void Populate(PowerPointDocument presentation,
        PowerPointBenchmarkFixture fixture) {
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(960),OfficeIMO.OfficeLength.Points(540));
        for (int index = 0; index < fixture.SlideCount; index++) {
            AddSlide(presentation, index, fixture.SlideCount);
        }
    }

    private static void AddSlide(PowerPointDocument presentation,
        int index, int slideCount) {
        PowerPointSlide slide = presentation.AddSlide();
        slide.BackgroundColor = index % 2 == 0 ? "F8FAFC" : "F1F5F9";

        PowerPointTextBox title = slide.AddTextBox(
            $"Operational review {index + 1}",global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(24),global::OfficeIMO.OfficeLength.Points(600),global::OfficeIMO.OfficeLength.Points(40));
        title.FontSize = global::OfficeIMO.OfficeLength.Points(24);
        title.Bold = true;
        title.Color = "0F172A";

        PowerPointTextBox subtitle = slide.AddTextBox(
            $"Slide {index + 1} of {slideCount} · deterministic benchmark corpus",            global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(72),global::OfficeIMO.OfficeLength.Points(700),global::OfficeIMO.OfficeLength.Points(28));
        subtitle.FontSize = global::OfficeIMO.OfficeLength.Points(12);
        subtitle.Color = "475569";

        string[] colors = { "DBEAFE", "DCFCE7", "FEF3C7", "FCE7F3" };
        for (int card = 0; card < 4; card++) {
            PowerPointAutoShape panel = slide.AddRectangle(
                OfficeIMO.OfficeLength.Points(40 + card * 220), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(190), OfficeIMO.OfficeLength.Points(72), $"Metric {card + 1}");
            panel.FillColor = colors[(card + index) % colors.Length];
            panel.OutlineColor = "CBD5E1";
            panel.OutlineWidthPoints = 1;
        }

        if (index % 3 == 0) AddTable(slide, index);
        else AddDetailPanels(slide, index);

        if (index % 5 == 0) AddChart(slide, index);
        else {
            PowerPointTextBox narrative = slide.AddTextBox(
                "Measured work includes editable text, vector shapes, package serialization, and rendering.",                global::OfficeIMO.OfficeLength.Points(390),global::OfficeIMO.OfficeLength.Points(238),global::OfficeIMO.OfficeLength.Points(500),global::OfficeIMO.OfficeLength.Points(110));
            narrative.FontSize = global::OfficeIMO.OfficeLength.Points(18);
            narrative.Color = "1E293B";
        }

        PowerPointTextBox footer = slide.AddTextBox(
            "OfficeIMO.PowerPoint performance corpus",global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(500),global::OfficeIMO.OfficeLength.Points(420),global::OfficeIMO.OfficeLength.Points(20));
        footer.FontSize = global::OfficeIMO.OfficeLength.Points(9);
        footer.Color = "64748B";
    }

    private static void AddTable(PowerPointSlide slide, int index) {
        PowerPointTable table = slide.AddTable(4,3,global::OfficeIMO.OfficeLength.Points(40),global::OfficeIMO.OfficeLength.Points(224),global::OfficeIMO.OfficeLength.Points(300),global::OfficeIMO.OfficeLength.Points(220));
        string[,] values = {
            { "Metric", "Current", "Target" },
            { "Quality", (92 + index % 7).ToString(), "98" },
            { "Coverage", (80 + index % 15).ToString(), "95" },
            { "Latency", (24 + index % 9).ToString(), "20" }
        };
        for (int row = 0; row < 4; row++) {
            for (int column = 0; column < 3; column++) {
                PowerPointTableCell cell = table.GetCell(row, column);
                cell.Text = values[row, column];
                if (row == 0) {
                    cell.Bold = true;
                    cell.FillColor = "DBEAFE";
                }
            }
        }
    }

    private static void AddDetailPanels(PowerPointSlide slide, int index) {
        for (int row = 0; row < 3; row++) {
            PowerPointAutoShape panel = slide.AddRectangle(
                OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeLength.Points(224 + row * 72), OfficeIMO.OfficeLength.Points(300), OfficeIMO.OfficeLength.Points(54), $"Detail {row + 1}");
            panel.FillColor = row % 2 == 0 ? "FFFFFF" : "F8FAFC";
            panel.OutlineColor = "CBD5E1";
            PowerPointTextBox value = slide.AddTextBox(
                $"Workstream {row + 1}: checkpoint {index + row + 1}",                global::OfficeIMO.OfficeLength.Points(56),global::OfficeIMO.OfficeLength.Points(239 + row * 72),global::OfficeIMO.OfficeLength.Points(260),global::OfficeIMO.OfficeLength.Points(24));
            value.FontSize = global::OfficeIMO.OfficeLength.Points(12);
            value.Color = "334155";
        }
    }

    private static void AddChart(PowerPointSlide slide, int index) {
        var data = new OfficeChartData(
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new[] {
                new OfficeSeries("Actual", new[] {
                    12D + index, 18D + index, 24D + index, 30D + index
                }),
                new OfficeSeries("Target", new[] {
                    15D + index, 20D + index, 26D + index, 32D + index
                })
            });
        slide.AddChart(OfficeChartType.BarClustered, data, new OfficeIMO.PowerPoint.PowerPointChartOptions { Left = global::OfficeIMO.OfficeLength.Points(390), Top = global::OfficeIMO.OfficeLength.Points(214), Width = global::OfficeIMO.OfficeLength.Points(500), Height = global::OfficeIMO.OfficeLength.Points(260) });
    }
}
