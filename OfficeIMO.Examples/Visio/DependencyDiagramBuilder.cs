using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Examples.Visio {
    public static class DependencyDiagramBuilder {
        public static void Example_DependencyDiagramBuilder(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Dependency diagram builder");
            string filePath = Path.Combine(folderPath, "Dependency Diagram Builder.vsdx");

            using var nativeDocument1 = VisioDocument.Create(filePath);
            nativeDocument1.DependencyDiagram("Service Dependencies", diagram => diagram
                    .Theme(VisioStyleTheme.Fluent())
                    .External("users", "Users")
                    .Component("web", "Web App")
                    .Component("api", "API")
                    .Decision("policy", "Policy")
                    .Data("database", "Database")
                    .Data("queue", "Queue")
                    .DependsOn("users", "web", "HTTPS")
                    .DependsOn("web", "api")
                    .ControlDependency("api", "policy", "Authorize")
                    .DataDependency("api", "database", "SQL")
                    .DataDependency("api", "queue", "Events"));
            nativeDocument1.EnsureVisualQuality(new VisioDiagramQualityOptions {
                    CheckConnectorShapeIntersections = false,
                    CheckConnectorLabelShapeOverlaps = false
                });
            nativeDocument1.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
