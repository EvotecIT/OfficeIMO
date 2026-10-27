using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class Charts
    {
        // Demo: palettes + full-width sizing helpers
        public static void Example_Charts_PaletteAndSizing(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating chart palette & sizing demo");
            string filePath = System.IO.Path.Combine(folderPath, "Charts.PaletteAndSizing.docx");
            using var doc = WordDocument.Create(filePath);
            // A quick intro
            doc.AddParagraph("Chart Palette & Sizing Showcase").SetStyle(WordParagraphStyles.Heading1);
            doc.AddParagraph("This example demonstrates the new chart helpers: ApplyPalette(...) and FitToPageContentWidth(...). ");
            // Build some sample data
            var categories = new List<string>
            {
                "Q1",
                "Q2",
                "Q3",
                "Q4"
            };
            // 1) Pie chart with semantic categories
            doc.AddParagraph("Pie (Professional palette, semantic outcomes)").SetStyle(WordParagraphStyles.Heading2);
            var pie = doc.AddChart("Rules outcome", false, 600, 320);
            pie.AddPie("Passed", 42);
            pie.AddPie("Failed", 30);
            pie.AddPie("Skipped", 5);
            pie.AddLegend(OfficeChartLegendPosition.Right);
            pie.ApplyPalette(WordChart.WordChartPalette.Professional, semanticOutcomes: true, applyToPies: true, applyToSeries: false).SetWidthToPageContent(1.0, 320);
            // 2) Bar chart (ColorBlindSafe palette)
            doc.AddParagraph("Bar (ColorBlindSafe palette)").SetStyle(WordParagraphStyles.Heading2);
            var bar = doc.AddChart("Quarterly revenue", false, 600, 320);
            bar.AddCategories(categories);
            bar.AddBar("EMEA", new List<int> { 10, 13, 16, 18 }, OfficeIMO.OfficeColor.Black); // initial color is overridden by palette
            bar.AddBar("APAC", new List<int> { 9, 12, 14, 20 }, OfficeIMO.OfficeColor.Black);
            bar.AddBar("AMER", new List<int> { 8, 15, 17, 19 }, OfficeIMO.OfficeColor.Black);
            bar.ApplyPalette(WordChart.WordChartPalette.ColorBlindSafe).SetWidthToPageContent(1.0, 320);
            // 3) Line chart (Monochrome palette), then override 1 series
            doc.AddParagraph("Line (MonochromeGray palette) with one override").SetStyle(WordParagraphStyles.Heading2);
            var line = doc.AddChart("KPIs", false, 600, 320);
            line.AddChartAxisX(categories);
            line.AddLine("Throughput", new List<int> { 100, 140, 180, 220 }, OfficeIMO.OfficeColor.Black);
            line.AddLine("Latency", new List<int> { 80, 60, 70, 50 }, OfficeIMO.OfficeColor.Black);
            line.ApplyPalette(WordChart.WordChartPalette.MonochromeGray).SetSeriesColor(1, OfficeIMO.OfficeColor.ParseHex("#d63939")) // emphasize Latency (series index 1)
            .SetWidthToPageContent(1.0, 320);
            // 4) Soft palette with no semantics
            doc.AddParagraph("Area (Soft palette, no semantics)").SetStyle(WordParagraphStyles.Heading2);
            var area = doc.AddChart("Utilization", false, 600, 320);
            area.AddCategories(categories);
            area.AddArea("CPU", new List<int> { 30, 45, 55, 35 }, OfficeIMO.OfficeColor.Black);
            area.AddArea("Memory", new List<int> { 40, 50, 60, 70 }, OfficeIMO.OfficeColor.Black);
            area.ApplyPalette(WordChart.WordChartPalette.Soft, semanticOutcomes: false).SetWidthToPageContent(1.0, 320);
            // Save/Open
            try
            {
                doc.Save(filePath);
            }
            catch
            {
                doc.Save();
            }

            if (openWord)
                doc.OpenInApplication(filePath);
        }
    }
}