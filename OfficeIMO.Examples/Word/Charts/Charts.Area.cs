using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Charts {
        public static void Example_AreaChart(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with an area chart");
            string filePath = System.IO.Path.Combine(folderPath, "AreaChart.docx");
            using WordDocument document = WordDocument.Create(filePath);
            List<string> categories = new() { "Food", "Housing", "Mix", "Data" };
            var areaChart = document.AddChart("Area chart");
            areaChart.AddCategories(categories);
            areaChart.AddArea("Brazil", new List<int> { 100, 1, 18, 230 }, Color.Brown);
            areaChart.AddArea("Poland", new List<int> { 13, 20, 230, 150 }, Color.Green);
            areaChart.AddArea("USA", new List<int> { 10, 305, 18, 23 }, Color.AliceBlue);
            areaChart.AddLegend(LegendPositionValues.Top);
            document.Save(openWord);
        }
    }
}
