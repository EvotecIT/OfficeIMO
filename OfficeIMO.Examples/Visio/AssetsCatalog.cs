using System;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates listing masters in a VSDX and selectively learning/extracting them.
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

            // Learn a supported subset into a new document, then use them by NameU.
            // The sample VSDX is an inspection fixture, not a runtime template.
            string filePath = Path.Combine(folderPath, "MasterCatalog Demo.vsdx");
            var doc = VisioDocument.Create(filePath);
            doc.UseMastersByDefault = true;
            doc.LearnMastersFromVsdx(sample, new[] { "Ellipse", "Diamond", "Triangle", "Rectangle" });
            var page = doc.AddPage("Assets Demo", OfficeLength.Centimeters(29.7), OfficeLength.Centimeters(21.0), scaleUnit: VisioMeasurementUnit.Centimeters);
            page.Shapes.Add(new VisioShape("1") { NameU = "Ellipse", PinX = OfficeLength.Inches(6), PinY = OfficeLength.Inches(12), Width = OfficeLength.Inches(4), Height = OfficeLength.Inches(3), Text = "Ellipse" });
            page.Shapes.Add(new VisioShape("2") { NameU = "Diamond", PinX = OfficeLength.Inches(13), PinY = OfficeLength.Inches(12), Width = OfficeLength.Inches(3.5), Height = OfficeLength.Inches(3), Text = "Diamond" });
            page.Shapes.Add(new VisioShape("3") { NameU = "Triangle", PinX = OfficeLength.Inches(6), PinY = OfficeLength.Inches(7), Width = OfficeLength.Inches(4), Height = OfficeLength.Inches(3), Text = "Triangle" });
            page.Shapes.Add(new VisioShape("4") { NameU = "Rectangle", PinX = OfficeLength.Inches(13), PinY = OfficeLength.Inches(7), Width = OfficeLength.Inches(4), Height = OfficeLength.Inches(2.5), Text = "Rectangle" });
            doc.Save();
        }
    }
}

