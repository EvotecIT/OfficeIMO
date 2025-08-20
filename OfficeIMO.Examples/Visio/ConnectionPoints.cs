using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates creating shapes with connection points and connecting them.
    /// </summary>
    public static class ConnectionPoints {
        public static void Example_ConnectionPoints(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Connection points");
            string filePath = Path.Combine(folderPath, "Connection Points.vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");

            VisioShape left = new("1", 2, 2, 2, 2, "Left");
            left.ConnectionPoints.Add(new VisioConnectionPoint(2, 1, 1, 0));
            page.Shapes.Add(left);

            VisioShape right = new("2", 6, 2, 2, 2, "Right");
            right.ConnectionPoints.Add(new VisioConnectionPoint(0, 1, -1, 0));
            page.Shapes.Add(right);

            VisioConnector connector = new(left, right) {
                FromConnectionPoint = left.ConnectionPoints[0],
                ToConnectionPoint = right.ConnectionPoints[0]
            };
            page.Connectors.Add(connector);

            document.Save(filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
