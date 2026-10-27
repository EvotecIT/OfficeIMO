using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Word {
    internal static partial class Charts {
        public static void Example_BarChart(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a bar chart");
            string filePath = System.IO.Path.Combine(folderPath, "BarChart.docx");
            using WordDocument document = WordDocument.Create(filePath);
            List<string> categories = new() { "Food", "Housing", "Mix", "Data" };
            var barChart = document.AddChart("Bar chart");
            barChart.AddCategories(categories);
            barChart.AddBar("USA", new List<int> { 10, 35, 18, 23 }, Color.Parse("#3F64D7"));
            barChart.AddBar("Brazil", new List<int> { 15, 30, 8, 18 }, Color.Brown);
            barChart.AddBar("Poland", new List<int> { 13, 20, 230, 150 }, Color.Green);
            barChart.BarGrouping = WordChartBarGrouping.Clustered;
            barChart.BarDirection = WordChartBarDirection.Column;
            barChart.AddLegend(OfficeChartLegendPosition.Top);
            document.Save();

            var valid = document.ValidateDocument();
            if (valid.Count > 0) {
                Console.WriteLine("Document has validation errors:");
                foreach (var error in valid) {
                    Console.WriteLine(error.Id + ": " + error.Description);
                }
            } else {
                Console.WriteLine("Document is valid.");
            }

            if (openWord) document.OpenInApplication();
        }
    }
}
