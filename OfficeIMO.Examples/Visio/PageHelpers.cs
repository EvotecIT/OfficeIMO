using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates page helpers for size, grid, shapes, and connectors.
    /// </summary>
    public static class PageHelpers {
        public static void Example_PageHelpers(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Page helpers");
            string filePath = Path.Combine(folderPath, "Page Helpers.vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            page.Size(11, 8.5).Grid(visible: true, snap: true);

            VisioMaster master = new("1", "Rectangle", new VisioShape("1"));
            VisioShape from = page.AddShape("1", master, 1, 1, 2, 1, "Start");
            VisioShape to = page.AddShape("2", master, 4, 1, 2, 1, "End");
            page.AddConnector("3", from, to, ConnectorKind.Straight);

            document.Save(filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
