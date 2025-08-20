using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates a connector with explicit styles.
    /// </summary>
    public static class ConnectorStyles {
        public static void Example_ConnectorStyles(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Connector with styles");
            string filePath = Path.Combine(folderPath, "Connector Styles.vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            VisioShape start = new("1", 1, 1, 2, 1, "Start") { NameU = "Rectangle" };
            VisioShape end = new("2", 4, 1, 2, 1, "End") { NameU = "Rectangle" };
            page.Shapes.Add(start);
            page.Shapes.Add(end);
            page.Connectors.Add(new VisioConnector(start, end) { EndArrow = 13 });
            document.Save(filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
