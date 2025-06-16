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
            document.Save(openWord);
        }
    }
}
