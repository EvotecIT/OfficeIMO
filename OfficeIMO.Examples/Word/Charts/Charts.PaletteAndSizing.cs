using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class Charts {
        // Demo: palettes + full-width sizing helpers
        public static void Example_Charts_PaletteAndSizing(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating chart palette & sizing demo");
            string filePath = System.IO.Path.Combine(folderPath, "Charts.PaletteAndSizing.docx");

            using var doc = WordDocument.Create(filePath);
            var fluent = new WordFluentDocument(doc);

            // A quick intro
            fluent.H1("Chart Palette & Sizing Showcase");
            fluent.P("This example demonstrates the new chart helpers: ApplyPalette(...) and FitToPageContentWidth(...). ");

            // Build some sample data
            var categories = new List<string> { "Q1", "Q2", "Q3", "Q4" };

            // 1) Pie chart with semantic categories
            fluent.H2("Pie (Professional palette, semantic outcomes)");
            var pie = doc.AddChart("Rules outcome", false, 600, 320);
            pie.AddPie("Passed", 42);
            pie.AddPie("Failed", 30);
            pie.AddPie("Skipped", 5);
            pie.AddLegend(DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues.Right);
            pie.ApplyPalette(WordChart.WordChartPalette.Professional, semanticOutcomes: true, applyToPies: true, applyToSeries: false)
               .SetWidthToPageContent(1.0, 320);

            // 2) Bar chart (ColorBlindSafe palette)
            fluent.H2("Bar (ColorBlindSafe palette)");
            var bar = doc.AddChart("Quarterly revenue", false, 600, 320);
            bar.AddCategories(categories);
            bar.AddBar("EMEA", new List<int> { 10, 13, 16, 18 }, SixLabors.ImageSharp.Color.Black); // initial color is overridden by palette
            bar.AddBar("APAC", new List<int> { 9, 12, 14, 20 }, SixLabors.ImageSharp.Color.Black);
            bar.AddBar("AMER", new List<int> { 8, 15, 17, 19 }, SixLabors.ImageSharp.Color.Black);
            bar.ApplyPalette(WordChart.WordChartPalette.ColorBlindSafe)
               .SetWidthToPageContent(1.0, 320);

            // 3) Line chart (Monochrome palette), then override 1 series
            fluent.H2("Line (MonochromeGray palette) with one override");
            var line = doc.AddChart("KPIs", false, 600, 320);
            line.AddChartAxisX(categories);
            line.AddLine("Throughput", new List<int> { 100, 140, 180, 220 }, SixLabors.ImageSharp.Color.Black);
            line.AddLine("Latency", new List<int> { 80, 60, 70, 50 }, SixLabors.ImageSharp.Color.Black);
            line.ApplyPalette(WordChart.WordChartPalette.MonochromeGray)
                .SetSeriesColor(1, SixLabors.ImageSharp.Color.ParseHex("#d63939")) // emphasize Latency (series index 1)
                .SetWidthToPageContent(1.0, 320);

            // 4) Soft palette with no semantics
            fluent.H2("Area (Soft palette, no semantics)");
            var area = doc.AddChart("Utilization", false, 600, 320);
            area.AddCategories(categories);
            area.AddArea("CPU", new List<int> { 30, 45, 55, 35 }, SixLabors.ImageSharp.Color.Black);
            area.AddArea("Memory", new List<int> { 40, 50, 60, 70 }, SixLabors.ImageSharp.Color.Black);
            area.ApplyPalette(WordChart.WordChartPalette.Soft, semanticOutcomes: false)
                .SetWidthToPageContent(1.0, 320);

            // Save/Open
            try { doc.Save(filePath, openWord); } catch { doc.Save(); }
        }
    }
}
