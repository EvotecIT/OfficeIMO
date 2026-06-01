using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class DataDrivenNetworkSegmentation {
        public static void Example_DataDrivenNetworkSegmentation(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Data-driven network segmentation diagram");
            string filePath = Path.Combine(folderPath, "Data Driven Network Segmentation Diagram.vsdx");

            VisioDocument document = VisioGallery.CreateNetworkSegmentationDiagram(filePath);
            document.Save();

            IReadOnlyList<string> issues = VisioValidator.Validate(filePath);
            if (issues.Count > 0) {
                throw new InvalidOperationException("Generated data-driven network segmentation diagram failed package validation:" + Environment.NewLine + string.Join(Environment.NewLine, issues));
            }

            document.EnsureVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = true,
                CheckConnectorShapeIntersections = true,
                CheckConnectorLabels = true
            });

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
