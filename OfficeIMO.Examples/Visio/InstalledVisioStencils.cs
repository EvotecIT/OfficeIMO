using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class InstalledVisioStencils {
        public static bool Example_InstalledVisioStencils(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Installed Visio stencil packages");
            IReadOnlyList<string> installedPackages = VisioStencilPackageCatalog.DiscoverInstalledVisioPackages();
            string[] selectedPackages = SelectInstalledPackages(installedPackages);
            if (selectedPackages.Length == 0) {
                Console.WriteLine("    Skipped: no installed Visio .vssx/.vstx packages were found.");
                return false;
            }

            VisioStencilCatalog catalog = VisioStencilPackageCatalog.LoadMany(selectedPackages, new VisioStencilPackageLoadOptions {
                CatalogName = "Installed Visio Stencils",
                IncludeUnsupportedMasters = true
            });

            string filePath = Path.Combine(folderPath, "Installed Visio Stencils - Hybrid Network.vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Installed Visio Stencils", 14, 8.5);
            page.Grid(visible: false, snap: true);

            AddHeader(page);
            AddZone(page, "Users", 1.75, 4.35, 2.4, 4.2, Color.FromRgb(245, 250, 255), Color.FromRgb(188, 214, 238));
            AddZone(page, "Network", 5.0, 4.35, 3.0, 4.2, Color.FromRgb(246, 252, 248), Color.FromRgb(179, 222, 190));
            AddZone(page, "Compute", 8.55, 4.35, 3.0, 4.2, Color.FromRgb(255, 250, 242), Color.FromRgb(237, 203, 153));
            AddZone(page, "Data", 12.05, 4.35, 2.25, 4.2, Color.FromRgb(250, 247, 255), Color.FromRgb(209, 194, 236));

            VisioStencilShape userDevice = Pick(catalog, "Laptop computer", "PC", "Tablet computer");
            VisioStencilShape gateway = Pick(catalog, "Virtual Network Gateways", "Azure VPN Gateway", "Load Balancers", "Front Doors");
            VisioStencilShape networkSecurity = Pick(catalog, "Network Security Groups", "Application Security Groups", "Firewall");
            VisioStencilShape app = Pick(catalog, "Function Apps", "App Services", "Web server");
            VisioStencilShape worker = Pick(catalog, "Virtual Machine", "Virtual Machines", "Server");
            VisioStencilShape data = Pick(catalog, "Azure SQL Database", "SQL databases", "Database server");

            VisioShape client = page.AddStencilShape(userDevice, "client", 1.75, 4.9, string.Empty);
            Label(page, "client-label", 1.75, 4.05, "Client");
            VisioShape vpn = page.AddStencilShape(gateway, "gateway", 4.25, 5.25, string.Empty);
            Label(page, "gateway-label", 4.25, 4.42, gateway.Name);
            VisioShape nsg = page.AddStencilShape(networkSecurity, "security", 5.8, 3.8, string.Empty);
            Label(page, "security-label", 5.8, 2.95, networkSecurity.Name);
            VisioShape function = page.AddStencilShape(app, "app", 8.0, 5.25, string.Empty);
            Label(page, "app-label", 8.0, 4.42, app.Name);
            VisioShape vm = page.AddStencilShape(worker, "worker", 9.55, 3.8, string.Empty);
            Label(page, "worker-label", 9.55, 2.95, worker.Name);
            VisioShape database = page.AddStencilShape(data, "database", 12.05, 4.6, string.Empty);
            Label(page, "database-label", 12.05, 3.75, data.Name);

            Connect(page, client, vpn, "HTTPS");
            Connect(page, vpn, function, "Ingress");
            Connect(page, vpn, nsg, "Policy");
            Connect(page, nsg, vm, "Private");
            Connect(page, function, database, "Reads/Writes");
            Connect(page, vm, database, "Batch");
            page.PolishDiagram(new VisioDiagramPolishOptions {
                FitToContent = false,
                ResizeShapesToText = false,
                ResizeConnectorLabelsToText = true,
                ResolveConnectorShapeIntersections = true,
                ResolveConnectorLabelOverlaps = true,
                ConnectorLabelMaxAttempts = 18,
                ConnectorLabelOptimizationPasses = 2
            });

            VisioShape note = page.AddTextBox("source-note", 7, 0.75, 11.5, 0.42,
                "Catalog composed from installed Microsoft Visio stencil packages: " + string.Join(", ", selectedPackages.Select(Path.GetFileName)));
            note.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 9.5,
                Color = Color.FromRgb(87, 96, 106),
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };

            document.Save();
            IReadOnlyList<string> issues = VisioValidator.Validate(filePath);
            if (issues.Count > 0) {
                throw new InvalidOperationException("Generated installed stencil example failed package validation:" + Environment.NewLine + string.Join(Environment.NewLine, issues));
            }

            Console.WriteLine("    Source packages: " + string.Join(", ", selectedPackages.Select(Path.GetFileName)));
            Console.WriteLine("    Output: " + filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }

            return true;
        }

        private static void AddHeader(VisioPage page) {
            page.AddTextBox("title", 7, 8.0, 12.5, 0.48, "Installed Visio stencil packages: hybrid network graph").TextStyle = new VisioTextStyle {
                FontFamily = "Aptos Display",
                Size = 21,
                Bold = true,
                Color = Color.FromRgb(32, 55, 75),
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };
            page.AddTextBox("subtitle", 7, 7.48, 10.8, 0.34, "Dependency-free .vssx/.vstx catalog discovery, source-aware auto-import, and graph-style placement").TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 10.5,
                Color = Color.FromRgb(85, 99, 113),
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };
        }

        private static void AddZone(VisioPage page, string title, double x, double y, double width, double height, Color fill, Color stroke) {
            VisioShape zone = page.AddRectangle(x, y, width, height, string.Empty);
            zone.FillColor = fill;
            zone.LineColor = stroke;
            zone.LineWeight = 0.012;
            zone.MarkAsBackgroundSurface();
            VisioShape label = page.AddTextBox("zone-" + title.Replace(" ", "-", StringComparison.OrdinalIgnoreCase), x, y + (height / 2) - 0.25, width - 0.25, 0.3, title);
            label.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 10,
                Bold = true,
                Color = Color.FromRgb(71, 85, 99),
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };
            label.MarkAsGeneratedDiagramAdornment();
        }

        private static void Label(VisioPage page, string id, double x, double y, string text) {
            VisioShape label = page.AddTextBox(id, x, y, 1.7, 0.36, text);
            label.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 9.5,
                Color = Color.FromRgb(33, 37, 41),
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };
            label.MarkAsGeneratedDiagramAdornment();
        }

        private static void Connect(VisioPage page, VisioShape from, VisioShape to, string label) {
            VisioConnector connector = page.AddConnector(from, to, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            connector.LineColor = Color.FromRgb(0, 120, 212);
            connector.LineWeight = 0.02;
            connector.EndArrow = EndArrow.Triangle;
            connector.Label = label;
            connector.LabelPlacement = VisioConnectorLabelPlacement.Along(0.5, 0, 0.16, 1.05, 0.28);
            connector.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 8.5,
                Color = Color.FromRgb(36, 79, 124),
                BackgroundColor = Color.FromRgb(255, 255, 255),
                BackgroundTransparency = 0
            };
        }

        private static VisioStencilShape Pick(VisioStencilCatalog catalog, params string[] queries) {
            foreach (string query in queries) {
                if (catalog.TryGet(query, out VisioStencilShape? exact) && exact != null) {
                    return exact;
                }

                VisioStencilShape? match = catalog.Search(query).FirstOrDefault();
                if (match != null) {
                    return match;
                }
            }

            throw new InvalidOperationException("None of the requested stencil shapes were found: " + string.Join(", ", queries));
        }

        private static string[] SelectInstalledPackages(IReadOnlyList<string> packages) {
            string[] preferredNames = {
                "COMPS_M.VSSX",
                "SERVER_M.VSSX",
                "AZURENETWORKING_M.VSSX",
                "AZURECOMPUTE_M.VSSX",
                "AZURECLOUD_M.VSSX",
                "AZUREDATABASES_M.VSSX"
            };

            List<string> selected = new();
            foreach (string preferredName in preferredNames) {
                string? match = packages.FirstOrDefault(path => string.Equals(Path.GetFileName(path), preferredName, StringComparison.OrdinalIgnoreCase));
                if (match != null) {
                    selected.Add(match);
                }
            }

            return selected.ToArray();
        }
    }
}
