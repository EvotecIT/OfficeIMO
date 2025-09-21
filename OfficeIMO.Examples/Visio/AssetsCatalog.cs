using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates listing masters in a VSDX and selectively importing/extracting them.
    /// </summary>
    internal static class AssetsCatalog {
        internal static void Example_ListAndExtractMasters(string folderPath, bool openVisio) {
            string assetsFolder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "VisioTemplates"));
            string sample = Path.Combine(assetsFolder, "DrawingWithLotsOfShapresAndArrows.vsdx");
            if (!File.Exists(sample)) return; // optional demo

            Console.WriteLine("[*] Visio - Master catalog from: " + Path.GetFileName(sample));
            var masters = VisioDocument.ListMastersIn(sample);
            foreach (var m in masters) Console.WriteLine($"  - {m.Id}: {m.NameU} ({m.Name})");

            // Extract a couple of masters to disk (for inspection or embedding later)
            string outDir = Path.Combine(folderPath, "ExtractedMasters");
            VisioAssets.ExtractMasters(sample, outDir, new[] { "Ellipse", "Diamond", "Triangle" });

            // Import a subset into a new document, then use them by NameU
            string filePath = Path.Combine(folderPath, "MasterCatalog Demo.vsdx");
            var doc = VisioDocument.Create(filePath);
            doc.UseMastersByDefault = true;
            doc.ImportMasters(sample, new[] { "Ellipse", "Diamond", "Triangle", "Rectangle" });
            var page = doc.AddPage("Assets Demo", 29.7, 21.0, VisioMeasurementUnit.Centimeters);
            page.Shapes.Add(new VisioShape("1") { NameU = "Ellipse", PinX = 6, PinY = 12, Width = 4, Height = 3, Text = "Ellipse" });
            page.Shapes.Add(new VisioShape("2") { NameU = "Diamond", PinX = 13, PinY = 12, Width = 3.5, Height = 3, Text = "Diamond" });
            page.Shapes.Add(new VisioShape("3") { NameU = "Triangle", PinX = 6, PinY = 7, Width = 4, Height = 3, Text = "Triangle" });
            page.Shapes.Add(new VisioShape("4") { NameU = "Rectangle", PinX = 13, PinY = 7, Width = 4, Height = 2.5, Text = "Rectangle" });
            doc.Save();
        }
    }
}

