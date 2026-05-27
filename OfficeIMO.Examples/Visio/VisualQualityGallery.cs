using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class VisualQualityGallery {
        public static void Example_VisualQualityGallery(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Visual quality gallery");
            string galleryPath = Path.Combine(folderPath, "Visio Gallery");
            VisioGalleryOptions options = new();
            if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_VISIO_DESKTOP_GALLERY"), "1", StringComparison.OrdinalIgnoreCase)) {
                options.ValidateWithVisioDesktop = true;
                options.DesktopValidationOptions = VisioDesktopValidationOptions.RoundTripWithSvg();
                options.DesktopValidationOptions.ExportDirectory = Path.Combine(galleryPath, "Visio Proof");
            }

            var results = VisioGallery.Create(galleryPath, options);

            foreach (VisioGalleryResult result in results) {
                Console.WriteLine($"  - {result.Name}: {(result.IsClean ? "clean" : "issues found")}");
                foreach (string issue in result.PackageIssues) {
                    Console.WriteLine($"    package: {issue}");
                }

                foreach (VisioDiagramQualityIssue issue in result.QualityIssues) {
                    Console.WriteLine($"    visual: {issue}");
                }

                if (result.DesktopValidation != null) {
                    foreach (string issue in result.DesktopValidation.Issues) {
                        Console.WriteLine($"    visio: {issue}");
                    }

                    foreach (string outputFile in result.DesktopValidation.OutputFiles) {
                        Console.WriteLine($"    proof: {outputFile}");
                    }
                }
            }

            string gatePath = Path.Combine(galleryPath, "Quality Gate Example.vsdx");
            VisioDocument gateDocument = VisioDocument.Create(gatePath);
            VisioPage gatePage = gateDocument.AddPage("Quality Gate", 6, 4);
            VisioShape source = gatePage.AddRectangle(1, 2, 0.8, 0.5, "Source");
            VisioShape target = gatePage.AddRectangle(5, 2, 0.8, 0.5, "Target");
            gatePage.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left).Label = "Flow";
            gateDocument.EnsureVisualQuality();
            gateDocument.Save();
            Console.WriteLine("  - Quality gate example: clean");

            if (openVisio && results.Count > 0) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(results.First().FilePath) { UseShellExecute = true });
            }
        }
    }
}
