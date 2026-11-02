using System;
using System.IO;
using OfficeIMO.Visio;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates connecting two colored rectangles with a styled connector.
    /// </summary>
    public static class ConnectRectangles {
        public static void Example_ConnectRectangles(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Connecting rectangles");
            string filePath = Path.Combine(folderPath, "Connect Rectangles.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            document.RequestRecalcOnOpen();
            VisioPage page = document.AddPage("Page-1", OfficeLength.Inches(8.5), OfficeLength.Inches(11));
            
            VisioShape start = new("1", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Start");
            
            VisioShape end = new("2", OfficeLength.Inches(4), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "End");
            
            page.Shapes.Add(start);
            page.Shapes.Add(end);
            
            var connector = new VisioConnector(start, end) {
                LineColor = Color.Blue,
                LineWeight = 0.02,
                Kind = ConnectorKind.RightAngle,
                BeginArrow = EndArrow.Arrow,
                EndArrow = EndArrow.Triangle,
                Label = "Start to End"
            };
            page.Connectors.Add(connector);
            
            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
