using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Examples.Visio {
    public static class MicrosoftIntegrationAzureStencils {
        private static readonly string[] PreferredPackageNames = {
            "Microsoft Integration Stencils.vssx",
            "MIS Azure Integration Services.vssx",
            "MIS Azure Stencils.vssx",
            "MIS Databases and Analytics Stencils.vssx",
            "MIS Infrastructure and Networking Stencils.vssx"
        };

        private static readonly IReadOnlyDictionary<string, string[]> RequiredStencilSelectors = new Dictionary<string, string[]> {
            ["client"] = new[] { "API", "Application", "3rd Party Integration" },
            ["apim"] = new[] { "API Management Services", "Azure: API Management Services" },
            ["servicebus"] = new[] { "Service Bus", "Azure: Service Bus" },
            ["eventgrid"] = new[] { "Event Grid", "Azure: Event Grid Topics", "Event Grid Topics" },
            ["logic"] = new[] { "Logic Apps", "Azure: Logic AppsLogic Apps", "Logic Apps Service" },
            ["function"] = new[] { "Function App", "Function" },
            ["sql"] = new[] { "SQL Databases", "Database", "Cloud Database" },
            ["storage"] = new[] { "Storage Accounts", "Azure: Storage Accounts", "Azure Blob Storage" },
            ["insights"] = new[] { "Application Insights", "Azure: Application Insights" }
        };

        public static void Example_MicrosoftIntegrationAzureStencils(string folderPath, bool openVisio, string stencilPackPathOrDirectory) {
            if (string.IsNullOrWhiteSpace(stencilPackPathOrDirectory)) throw new ArgumentException("Stencil pack path cannot be null or whitespace.", nameof(stencilPackPathOrDirectory));

            Console.WriteLine("[*] Visio - Microsoft Integration/Azure external stencil pack graph");
            string filePath = Path.Combine(folderPath, "Microsoft Integration Azure Stencil Graph.vsdx");
            string[] packagePaths = ResolvePackagePaths(stencilPackPathOrDirectory).ToArray();
            if (packagePaths.Length == 0) {
                throw new FileNotFoundException("No supported Visio stencil packages were found.", stencilPackPathOrDirectory);
            }

            VisioStencilCatalog catalog = VisioStencilPackageCatalog.LoadMany(packagePaths, new VisioStencilPackageLoadOptions {
                CatalogName = "Microsoft Integration and Azure Stencils",
                Category = "Microsoft Integration and Azure",
                IncludeUnsupportedMasters = true
            });

            Dictionary<string, VisioStencilShape> stencils = PickRequired(catalog, RequiredStencilSelectors).ToDictionary(
                pair => pair.Key,
                pair => AsDiagramIcon(pair.Value),
                StringComparer.Ordinal);

            VisioDocument.Create(filePath)
                .GraphDiagram("Order processing integration map", graph => graph
                    .Title()
                    .Theme(VisioStyleTheme.Technical())
                    .Layout(VisioGraphLayout.Layered)
                    .Direction(VisioGraphDirection.LeftToRight)
                    .PageSize(15.5, 7.6)
                    .Margins(0.9, 0.9, 0.9, 1.05)
                    .NodeSize(1.08, 0.82)
                    .Spacing(1.18, 1.18)
                    .StencilNode("client", "Partner API", stencils["client"])
                    .StencilNode("apim", "API Management", stencils["apim"])
                    .StencilNode("servicebus", "Service Bus", stencils["servicebus"])
                    .StencilNode("eventgrid", "Event Grid", stencils["eventgrid"])
                    .StencilNode("logic", "Logic App", stencils["logic"])
                    .StencilNode("function", "Function", stencils["function"])
                    .StencilNode("sql", "SQL Database", stencils["sql"])
                    .StencilNode("storage", "Blob Storage", stencils["storage"])
                    .StencilNode("insights", "App Insights", stencils["insights"])
                    .Root("client")
                    .Zone("edge", "API edge", "client", "apim")
                    .Zone("integration", "Integration workflow", "servicebus", "eventgrid", "logic", "function")
                    .Zone("data", "Operational data", "sql", "storage")
                    .Zone("ops", "Observability", "insights")
                    .Edge("client", "apim", "REST")
                    .ControlEdge("apim", "servicebus", "command")
                    .ControlEdge("servicebus", "logic", "workflow")
                    .ControlEdge("servicebus", "eventgrid", "event")
                    .Edge("eventgrid", "function", "trigger")
                    .DataEdge("logic", "sql", "upsert")
                    .DataEdge("function", "storage", "archive")
                    .Relationship("apim", "insights", "telemetry")
                    .Relationship("function", "insights", "traces"))
                .EnsureVisualQuality(new VisioDiagramQualityOptions {
                    CheckShapeOverlaps = false,
                    CheckConnectorShapeIntersections = false,
                    CheckConnectorLabelShapeOverlaps = false
                })
                .Save();

            IReadOnlyList<string> issues = VisioValidator.Validate(filePath);
            if (issues.Count > 0) {
                throw new InvalidOperationException("Generated Microsoft Integration/Azure stencil example failed package validation:" + Environment.NewLine + string.Join(Environment.NewLine, issues));
            }

            Console.WriteLine("    Source packages: " + string.Join(", ", packagePaths.Select(Path.GetFileName)));
            Console.WriteLine("    Imported masters: " + string.Join(", ", stencils.Values.Select(shape => shape.MasterNameU).Distinct(StringComparer.OrdinalIgnoreCase)));
            Console.WriteLine("    Output: " + filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }

        public static string? ResolveConfiguredPackPath(string[] args) {
            string? argument = GetArgumentValue(args, "--visio-integration-stencils")
                ?? GetArgumentValue(args, "--visio-microsoft-integration-stencils");
            if (!string.IsNullOrWhiteSpace(argument)) {
                return argument;
            }

            return Environment.GetEnvironmentVariable("OFFICEIMO_VISIO_INTEGRATION_STENCILS");
        }

        public static bool IsConfigured(string? pathOrDirectory) {
            if (string.IsNullOrWhiteSpace(pathOrDirectory)) {
                return false;
            }

            if (File.Exists(pathOrDirectory)) {
                string fullPath = Path.GetFullPath(pathOrDirectory);
                string directory = Path.GetDirectoryName(fullPath) ?? ".";
                bool supportedPackage = VisioStencilPackageCatalog.EnumeratePackageFiles(directory, recursive: false)
                    .Contains(fullPath, StringComparer.OrdinalIgnoreCase);
                return supportedPackage && HasRequiredStencils(new[] { fullPath });
            }

            if (!Directory.Exists(pathOrDirectory)) {
                return false;
            }

            IReadOnlyList<string> packages = VisioStencilPackageCatalog.EnumeratePackageFiles(pathOrDirectory, recursive: true);
            string[] selectedPackages = packages.Where(IsPreferredPackage).ToArray();
            return selectedPackages.Length > 0 && HasRequiredStencils(selectedPackages);
        }

        private static IEnumerable<string> ResolvePackagePaths(string pathOrDirectory) {
            if (File.Exists(pathOrDirectory)) {
                yield return Path.GetFullPath(pathOrDirectory);
                yield break;
            }

            if (!Directory.Exists(pathOrDirectory)) {
                throw new DirectoryNotFoundException("Visio stencil pack directory was not found: " + pathOrDirectory);
            }

            IReadOnlyList<string> packages = VisioStencilPackageCatalog.EnumeratePackageFiles(pathOrDirectory, recursive: true);
            List<string> selected = new();
            foreach (string preferredName in PreferredPackageNames) {
                string? package = packages.FirstOrDefault(path => string.Equals(Path.GetFileName(path), preferredName, StringComparison.OrdinalIgnoreCase));
                if (package != null) {
                    selected.Add(package);
                }
            }

            foreach (string package in selected) {
                yield return package;
            }
        }

        private static bool IsPreferredPackage(string packagePath) {
            return PreferredPackageNames.Contains(Path.GetFileName(packagePath), StringComparer.OrdinalIgnoreCase);
        }

        private static bool HasRequiredStencils(IEnumerable<string> packagePaths) {
            try {
                VisioStencilCatalog catalog = VisioStencilPackageCatalog.LoadMany(packagePaths, new VisioStencilPackageLoadOptions {
                    CatalogName = "Microsoft Integration and Azure Stencils",
                    Category = "Microsoft Integration and Azure",
                    IncludeUnsupportedMasters = true
                });
                foreach (string[] selectors in RequiredStencilSelectors.Values) {
                    if (!catalog.TryFindBest(selectors, out VisioStencilShape? stencil) || stencil == null) {
                        return false;
                    }
                }

                return true;
            } catch {
                return false;
            }
        }

        private static Dictionary<string, VisioStencilShape> PickRequired(VisioStencilCatalog catalog, IReadOnlyDictionary<string, string[]> selectors) {
            Dictionary<string, VisioStencilShape> selected = new(StringComparer.Ordinal);
            foreach (KeyValuePair<string, string[]> selector in selectors) {
                if (catalog.TryFindBest(selector.Value, out VisioStencilShape? stencil) && stencil != null) {
                    selected.Add(selector.Key, stencil);
                    continue;
                }

                throw new InvalidOperationException("Required stencil was not found for '" + selector.Key + "'. Tried: " + string.Join(", ", selector.Value));
            }

            return selected;
        }

        private static VisioStencilShape AsDiagramIcon(VisioStencilShape stencil) {
            double longestSide = Math.Max(stencil.DefaultWidth, stencil.DefaultHeight);
            if (longestSide <= 0.88D) {
                return stencil;
            }

            double scale = 0.88D / longestSide;
            return new VisioStencilShape(
                stencil.Id,
                stencil.Name,
                stencil.MasterNameU,
                stencil.Category,
                Math.Max(0.42D, stencil.DefaultWidth * scale),
                Math.Max(0.42D, stencil.DefaultHeight * scale),
                stencil.Keywords,
                stencil.Aliases,
                stencil.Tags,
                stencil.IconNameU,
                stencil.DefaultUnit,
                stencil.SourcePackagePath);
        }

        private static string? GetArgumentValue(string[] args, string name) {
            for (int i = 0; i < args.Length; i++) {
                if (!string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal)) {
                    return args[i + 1];
                }

                return null;
            }

            string prefix = name + "=";
            return args.FirstOrDefault(arg => arg.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))?.Substring(prefix.Length);
        }
    }
}
