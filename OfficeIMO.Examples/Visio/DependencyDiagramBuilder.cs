using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Examples.Visio {
    public static class DependencyDiagramBuilder {
        public static void Example_DependencyDiagramBuilder(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Dependency diagram builder");
            string filePath = Path.Combine(folderPath, "Dependency Diagram Builder.vsdx");

            VisioDocument.Create(filePath)
                .DependencyDiagram("Service Dependencies", diagram => diagram
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
                    .DataDependency("api", "queue", "Events"))
                .EnsureVisualQuality(new VisioDiagramQualityOptions {
                    CheckConnectorShapeIntersections = false,
                    CheckConnectorLabelShapeOverlaps = false
                })
                .Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
