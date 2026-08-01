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
        using (PowerPointPresentation presentation =
               PowerPointPresentation.Create(stream)) {
            Populate(presentation, fixture);
            presentation.Save();
        }
        return stream.ToArray();
    }

    internal static void Populate(PowerPointPresentation presentation,
        PowerPointBenchmarkFixture fixture) {
        presentation.SlideSize.SetSizePoints(960, 540);
        for (int index = 0; index < fixture.SlideCount; index++) {
            AddSlide(presentation, index, fixture.SlideCount);
        }
    }

    private static void AddSlide(PowerPointPresentation presentation,
        int index, int slideCount) {
        PowerPointSlide slide = presentation.AddSlide();
        slide.BackgroundColor = index % 2 == 0 ? "F8FAFC" : "F1F5F9";

        PowerPointTextBox title = slide.AddTextBoxPoints(
            $"Operational review {index + 1}", 40, 24, 600, 40);
        title.FontSize = 24;
        title.Bold = true;
        title.Color = "0F172A";

        PowerPointTextBox subtitle = slide.AddTextBoxPoints(
            $"Slide {index + 1} of {slideCount} · deterministic benchmark corpus",
            40, 72, 700, 28);
        subtitle.FontSize = 12;
        subtitle.Color = "475569";

        string[] colors = { "DBEAFE", "DCFCE7", "FEF3C7", "FCE7F3" };
        for (int card = 0; card < 4; card++) {
            PowerPointAutoShape panel = slide.AddRectanglePoints(
                40 + card * 220, 120, 190, 72, $"Metric {card + 1}");
            panel.FillColor = colors[(card + index) % colors.Length];
            panel.OutlineColor = "CBD5E1";
            panel.OutlineWidthPoints = 1;
        }

        if (index % 3 == 0) AddTable(slide, index);
        else AddDetailPanels(slide, index);

        if (index % 5 == 0) AddChart(slide, index);
        else {
            PowerPointTextBox narrative = slide.AddTextBoxPoints(
                "Measured work includes editable text, vector shapes, package serialization, and rendering.",
                390, 238, 500, 110);
            narrative.FontSize = 18;
            narrative.Color = "1E293B";
        }

        PowerPointTextBox footer = slide.AddTextBoxPoints(
            "OfficeIMO.PowerPoint performance corpus", 40, 500, 420, 20);
        footer.FontSize = 9;
        footer.Color = "64748B";
    }

    private static void AddTable(PowerPointSlide slide, int index) {
        PowerPointTable table = slide.AddTablePoints(4, 3, 40, 224, 300, 220);
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
            PowerPointAutoShape panel = slide.AddRectanglePoints(
                40, 224 + row * 72, 300, 54, $"Detail {row + 1}");
            panel.FillColor = row % 2 == 0 ? "FFFFFF" : "F8FAFC";
            panel.OutlineColor = "CBD5E1";
            PowerPointTextBox value = slide.AddTextBoxPoints(
                $"Workstream {row + 1}: checkpoint {index + row + 1}",
                56, 239 + row * 72, 260, 24);
            value.FontSize = 12;
            value.Color = "334155";
        }
    }

    private static void AddChart(PowerPointSlide slide, int index) {
        var data = new OfficeChartData(
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new[] {
                new OfficeChartSeries("Actual", new[] {
                    12D + index, 18D + index, 24D + index, 30D + index
                }),
                new OfficeChartSeries("Target", new[] {
                    15D + index, 20D + index, 26D + index, 32D + index
                })
            });
        PowerPointChart chart = slide.AddChartPoints(
            OfficeChartKind.BarClustered, data, 390, 214, 500, 260);
        chart.SetTitle("Quarterly trajectory");
    }
}
