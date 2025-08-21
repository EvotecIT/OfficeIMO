using System;
using System.IO;
using OfficeIMO.Visio;
using SixLabors.ImageSharp;

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
            VisioPage page = document.AddPage("Page-1", 8.5, 11);
            
            VisioShape start = new("1", 1, 1, 2, 1, "Start") { 
                NameU = "Rectangle",
                FillColor = Color.LightGreen,
                LineColor = Color.DarkGreen,
                LineWeight = 0.025
            };
            
            VisioShape end = new("2", 4, 1, 2, 1, "End") { 
                NameU = "Rectangle",
                FillColor = Color.LightCoral,
                LineColor = Color.DarkRed,
                LineWeight = 0.025
            };
            
            page.Shapes.Add(start);
            page.Shapes.Add(end);
            
            var connector = new VisioConnector(start, end) {
                LineColor = Color.Blue,
                LineWeight = 0.02,
                EndArrow = 1 // Add arrow at the end
            };
            page.Connectors.Add(connector);
            
            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
