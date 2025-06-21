using System.Collections.Generic;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Charts {
        public static void Example_RadarChart(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a radar chart");
            string filePath = System.IO.Path.Combine(folderPath, "RadarChart.docx");
            using WordDocument document = WordDocument.Create(filePath);
            List<string> categories = new() { "Food", "Housing", "Mix", "Data" };
            var radarChart = document.AddChart("Radar chart");
            radarChart.AddCategories(categories);
            radarChart.AddRadar("USA", new List<int> { 1, 5, 3, 2 }, Color.Blue);
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
