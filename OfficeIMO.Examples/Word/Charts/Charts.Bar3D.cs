using System.Collections.Generic;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Charts {
        public static void Example_Bar3DChart(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a 3-D bar chart");
            string filePath = System.IO.Path.Combine(folderPath, "Bar3DChart.docx");
            using WordDocument document = WordDocument.Create(filePath);
            List<string> categories = new() { "Food", "Housing", "Mix", "Data" };
            var bar3d = document.AddChart("Bar3D chart");
            bar3d.AddCategories(categories);
            bar3d.AddBar3D("USA", new List<int> { 5, 2, 3, 4 }, Color.DarkOrange);
            document.Save(openWord);
        }
    }
}
