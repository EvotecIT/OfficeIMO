using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Charts {
        public static void Example_ComboChart(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a combo chart");
            string filePath = System.IO.Path.Combine(folderPath, "ComboChart.docx");
            using WordDocument document = WordDocument.Create(filePath);
            List<string> categories = new() { "Food", "Housing", "Mix", "Data" };
            var comboChart = document.AddComboChart("Combo chart");
            comboChart.AddChartAxisX(categories);
            comboChart.AddBar("Sales", new List<int> { 10, 35, 18, 23 }, Color.Brown);
            comboChart.AddLine("Trend", new List<int> { 12, 30, 20, 25 }, Color.AliceBlue);
            comboChart.AddLegend(LegendPositionValues.Top);
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
