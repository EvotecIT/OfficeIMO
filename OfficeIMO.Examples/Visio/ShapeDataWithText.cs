using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates adding shape data along with text.
    /// </summary>
    public static class ShapeDataWithText {
        public static void Example_ShapeDataWithText(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Shape Data with Text");
            string filePath = Path.Combine(folderPath, "Shape Data with Text.vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            VisioShape shape = new("1", 2, 2, 3, 2, "A");
            shape.Data["Key"] = "Value";
            shape.Text = "Hello";
            page.Shapes.Add(shape);
            document.Save(filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
