using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates manipulating shape transform and line properties.
    /// </summary>
    public static class ShapeProperties {
        public static void Example_ShapeProperties(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Shape Properties");
            string filePath = Path.Combine(folderPath, "Shape Properties.vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            VisioShape shape = new("1", 2, 2, 3, 2, "A") {
                LineWeight = 0.02,
                LocPinX = 1.5,
                LocPinY = 1,
                Angle = 0.2,
            };
            page.Shapes.Add(shape);
            document.Save(filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
