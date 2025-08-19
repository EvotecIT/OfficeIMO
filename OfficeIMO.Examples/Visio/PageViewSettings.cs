using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates configuring page view settings and basic shape cells.
    /// </summary>
    public static class PageViewSettings {
        public static void Example_PageViewSettings(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Page view settings");
            string filePath = Path.Combine(folderPath, "Page view settings.vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            page.ViewScale = 1.5;
            page.ViewCenterX = 4;
            page.ViewCenterY = 5;
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Rectangle"));
            document.Save(filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

