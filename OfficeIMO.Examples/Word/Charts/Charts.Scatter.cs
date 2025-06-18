using System.Collections.Generic;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Charts {
        public static void Example_ScatterChart(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a scatter chart");
            string filePath = System.IO.Path.Combine(folderPath, "ScatterChart.docx");
            using WordDocument document = WordDocument.Create(filePath);
            var scatterChart = document.AddChart("Scatter chart");
            scatterChart.AddScatter("Data", new List<double> { 1, 2, 3 }, new List<double> { 3, 2, 1 }, Color.Red);

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
