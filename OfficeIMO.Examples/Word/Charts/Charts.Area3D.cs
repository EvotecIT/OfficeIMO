using System;
using System.Collections.Generic;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Charts {
        public static void Example_Area3DChart(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a 3-D area chart");
            string filePath = System.IO.Path.Combine(folderPath, "Area3DChart.docx");
            using WordDocument document = WordDocument.Create(filePath);
            List<string> categories = new() { "Food", "Housing", "Mix", "Data" };
            var area3d = document.AddChart("Area3D chart");
            area3d.AddCategories(categories);
            area3d.AddArea3D("USA", new List<int> { 5, 2, 3, 4 }, Color.DarkBlue);
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
