using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates connecting two rectangles with an orthogonal connector.
    /// </summary>
    public static class ConnectRectangles {
        public static void Example_ConnectRectangles(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Connecting rectangles");
            string filePath = Path.Combine(folderPath, "Connect Rectangles.vsdx");

            VisioDocument document = new();
            document.RequestRecalcOnOpen();
            VisioPage page = document.AddPage("Page-1");
            VisioShape start = new("1", 1, 1, 2, 1, "Start");
            VisioShape end = new("2", 4, 1, 2, 1, "End");
            page.Shapes.Add(start);
            page.Shapes.Add(end);
            page.Connectors.Add(new VisioConnector(start, end));
            document.Save(filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
