using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates adding shape data.
    /// </summary>
    public static class ShapeData {
        public static void Example_ShapeData(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Shape Data");
            string filePath = Path.Combine(folderPath, "Shape Data.vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            VisioShape shape = new("1", 2, 2, 3, 2, "A");
            shape.Data["Key"] = "Value";
            page.Shapes.Add(shape);
            document.Save(filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
