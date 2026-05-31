using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class DataDrivenApplicationGraph {
        public static void Example_DataDrivenApplicationGraph(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Data-driven application dependency graph");
            string filePath = Path.Combine(folderPath, "Data Driven Application Dependency Graph.vsdx");

            VisioDocument document = VisioGallery.CreateApplicationDependencyGraph(filePath);
            document.Save();

            IReadOnlyList<string> issues = VisioValidator.Validate(filePath);
            if (issues.Count > 0) {
                throw new InvalidOperationException("Generated data-driven application dependency graph failed package validation:" + Environment.NewLine + string.Join(Environment.NewLine, issues));
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
