using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates creating a rectangle with explicit styles so it remains visible.
    /// </summary>
    public static class RectangleStyles {
        public static void Example_RectangleStyles(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Rectangle with styles");
            string filePath = Path.Combine(folderPath, "Rectangle Styles.vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Rect"));
            document.Save(filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
