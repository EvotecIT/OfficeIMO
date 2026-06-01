using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class DataDrivenKubernetesGraph {
        public static void Example_DataDrivenKubernetesGraph(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Data-driven Kubernetes service-mesh graph");
            string filePath = Path.Combine(folderPath, "Data Driven Kubernetes Service Mesh Graph.vsdx");

            VisioDocument document = VisioGallery.CreateKubernetesServiceMeshGraph(filePath);
            document.Save();

            IReadOnlyList<string> issues = VisioValidator.Validate(filePath);
            if (issues.Count > 0) {
                throw new InvalidOperationException("Generated data-driven Kubernetes graph failed package validation:" + Environment.NewLine + string.Join(Environment.NewLine, issues));
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
