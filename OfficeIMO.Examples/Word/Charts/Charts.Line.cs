using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Charts {
        public static void Example_LineChart(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a line chart");
            string filePath = System.IO.Path.Combine(folderPath, "LineChart.docx");
            using WordDocument document = WordDocument.Create(filePath);
            List<string> categories = new() { "Food", "Housing", "Mix", "Data" };
            var lineChart = document.AddChart("Line chart");
            lineChart.AddChartAxisX(categories);
            lineChart.AddLine("USA", new List<int> { 10, 35, 18, 23 }, Color.AliceBlue);
            lineChart.AddLine("Brazil", new List<int> { 10, 35, 300, 18 }, Color.Brown);
            lineChart.AddLine("Poland", new List<int> { 13, 20, 230, 150 }, Color.Green);
            document.Save(false);

            var valid = document.ValidateDocument();
            if (valid.Count > 0) {
                Console.WriteLine("Document has validation errors:");
                foreach (var error in valid) {
                    Console.WriteLine(error.Id + ": " + error.Description);
                }
            } else {
                Console.WriteLine("Document is valid.");
            }

            document.Open(openWord);
        }
    }
}
