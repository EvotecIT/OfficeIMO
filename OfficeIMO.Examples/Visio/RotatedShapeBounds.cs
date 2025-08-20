using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates computing bounds for a rotated shape.
    /// </summary>
    public static class RotatedShapeBounds {
        public static void Example_RotatedShapeBounds(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Rotated Shape Bounds");
            string filePath = Path.Combine(folderPath, "Rotated Shape Bounds.vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            VisioShape shape = new("1", 2, 2, 2, 1, "A") {
                Angle = Math.PI / 4,
            };
            page.Shapes.Add(shape);

            (double left, double bottom, double right, double top) = shape.GetBounds();
            Console.WriteLine($"Bounds: L={left}, B={bottom}, R={right}, T={top}");

            document.Save(filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

