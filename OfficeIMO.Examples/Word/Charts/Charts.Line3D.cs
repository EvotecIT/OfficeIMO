using System;
using System.Collections.Generic;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Charts {
        public static void Example_Line3DChart(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a 3-D line chart");
            string filePath = System.IO.Path.Combine(folderPath, "Line3DChart.docx");
            using WordDocument document = WordDocument.Create(filePath);
            List<string> categories = new() { "Food", "Housing", "Mix", "Data" };
            var line3d = document.AddChart("Line3D chart");
            line3d.AddChartAxisX(categories);
            line3d.AddLine3D("USA", new List<int> { 5, 2, 3, 4 }, Color.Purple);
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
