using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Examples.Visio {
    public static class GraphDiagramBuilder {
        public static void Example_GraphDiagramBuilder(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Graph diagram builder");
            string filePath = Path.Combine(folderPath, "Graph Diagram Builder.vsdx");
            VisioStencilCatalog? installedCatalog = TryLoadInstalledAzureCatalog();

            VisioDocument.Create(filePath)
                .GraphDiagram("Event-driven operations graph", graph => {
                    graph
                        .Title()
                        .Theme(VisioStyleTheme.Technical())
                        .Layout(VisioGraphLayout.Layered)
                        .Direction(VisioGraphDirection.LeftToRight)
                        .PageSize(13.5, 5.7)
                        .Margins(0.8, 0.85, 0.8, 0.7)
                        .NodeSize(1.25, 0.72)
                        .Spacing(0.68, 0.86);

                    AddNode(graph, installedCatalog, "users", "Users", VisioGraphNodeKind.External, "Users", "Laptop computer", "PC");
                    AddNode(graph, installedCatalog, "gateway", "API gateway", VisioGraphNodeKind.Emphasis, "API Management Services", "Application Gateway", "Front Doors");
                    AddNode(graph, installedCatalog, "events", "Events", VisioGraphNodeKind.Process, "Event Grid", "Service Bus", "Queues");
                    AddNode(graph, installedCatalog, "function", "Function", VisioGraphNodeKind.Process, "Function Apps", "Azure Functions");
                    AddNode(graph, installedCatalog, "worker", "Worker", VisioGraphNodeKind.Process, "Virtual Machine", "App Services", "Web server");
                    AddNode(graph, installedCatalog, "sql", "SQL", VisioGraphNodeKind.Data, "Azure SQL Database", "SQL databases", "Database server");
                    AddNode(graph, installedCatalog, "monitor", "Monitor", VisioGraphNodeKind.External, "Azure Monitor", "Monitor", "Application Insights");
                    AddNode(graph, installedCatalog, "batch", "Batch", VisioGraphNodeKind.Emphasis, "Automation Accounts", "Logic Apps", "Runbooks");

                    graph
                        .Root("users")
                        .Zone("edge", "Edge", "users", "gateway")
                        .Zone("runtime", "Runtime", "events", "function", "worker", "batch", "monitor")
                        .Zone("data", "Data", "sql")
                        .Edge("users", "gateway", "HTTPS")
                        .ControlEdge("gateway", "events", "publish")
                        .ControlEdge("events", "function", "trigger")
                        .Edge("function", "worker", "dispatch")
                        .ControlEdge("worker", "batch", "schedule")
                        .DataEdge("worker", "sql", "write")
                        .DataEdge("sql", "function", "read model")
                        .Relationship("monitor", "function", "metrics")
                        .Relationship("monitor", "worker", "logs")
                        .EmphasisEdge("batch", "sql", "reconcile");
                })
                .EnsureVisualQuality(new VisioDiagramQualityOptions {
                    CheckShapeOverlaps = false,
                    CheckConnectorShapeIntersections = false,
                    CheckConnectorLabelShapeOverlaps = false
                })
                .Save();

            IReadOnlyList<string> issues = VisioValidator.Validate(filePath);
            if (issues.Count > 0) {
                throw new InvalidOperationException("Generated graph example failed package validation:" + Environment.NewLine + string.Join(Environment.NewLine, issues));
            }

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }

        private static VisioStencilCatalog? TryLoadInstalledAzureCatalog() {
            string[] selectedPackages = VisioStencilPackageCatalog.DiscoverInstalledVisioPackages()
                .Where(path => Path.GetFileName(path).StartsWith("AZURE", StringComparison.OrdinalIgnoreCase) ||
                               string.Equals(Path.GetFileName(path), "COMPS_M.VSSX", StringComparison.OrdinalIgnoreCase) ||
                               string.Equals(Path.GetFileName(path), "SERVER_M.VSSX", StringComparison.OrdinalIgnoreCase))
                .Take(12)
                .ToArray();

            if (selectedPackages.Length == 0) {
                return null;
            }

            return VisioStencilPackageCatalog.LoadMany(selectedPackages, new VisioStencilPackageLoadOptions {
                CatalogName = "Installed Visio Graph Stencils",
                IncludeUnsupportedMasters = true
            });
        }

        private static void AddNode(VisioGraphDiagramBuilder graph, VisioStencilCatalog? catalog, string id, string text, VisioGraphNodeKind fallbackKind, params string[] queries) {
            if (catalog != null && catalog.TryFindBest(queries, out VisioStencilShape? stencil) && stencil != null) {
                graph.StencilNode(id, text, catalog, queries);
                return;
            }

            graph.Node(id, text, fallbackKind);
        }
    }
}
