using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class ExternalStencilPack {
        public static void Example_ExternalStencilPack(string folderPath, bool openVisio, string stencilPackPath) {
            if (string.IsNullOrWhiteSpace(stencilPackPath)) throw new ArgumentException("Stencil pack path cannot be null or whitespace.", nameof(stencilPackPath));
            if (!File.Exists(stencilPackPath)) throw new FileNotFoundException("Stencil pack was not found.", stencilPackPath);

            Console.WriteLine("[*] Visio - External VSSX stencil pack");
            string filePath = Path.Combine(folderPath, "External VSSX Stencil Pack.vsdx");

            VisioStencilCatalog catalog = VisioStencilPackageCatalog.Load(stencilPackPath, new VisioStencilPackageLoadOptions {
                CatalogName = "External Stencils",
                Category = "External",
                IncludeUnsupportedMasters = true
            });

            IReadOnlyList<VisioStencilShape> selected = SelectStencils(catalog,
                "API Management Services",
                "Logic Apps",
                "Service Bus",
                "Event Grid",
                "Function App",
                "Storage accounts");
            if (selected.Count < 4) {
                selected = catalog.Shapes.Take(6).ToList();
            }

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("External Stencil Pack", 14, 8.5);
            page.Grid(visible: false, snap: true);

            VisioShape frame = page.AddRectangle(7, 4.25, 13.2, 6.7, string.Empty);
            frame.FillColor = Color.FromRgb(247, 250, 252);
            frame.LineColor = Color.FromRgb(218, 230, 242);
            frame.LineWeight = 0.012;

            VisioShape lane = page.AddRectangle(7, 5.75, 12.4, 2.35, string.Empty);
            lane.FillColor = Color.FromRgb(255, 255, 255);
            lane.LineColor = Color.FromRgb(205, 220, 236);
            lane.LineWeight = 0.01;

            VisioTextStyle titleStyle = new() {
                FontFamily = "Aptos Display",
                Size = 22,
                Bold = true,
                Color = Color.FromRgb(32, 55, 75),
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };
            page.AddTextBox("title", 7, 8.0, 12.5, 0.5, "External VSSX stencil pack: real imported masters").TextStyle = titleStyle;

            page.AddTextBox("subtitle", 7, 7.45, 10.8, 0.34, "Microsoft Integration and Azure stencil artwork imported from a .vssx package").TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 11,
                Color = Color.FromRgb(85, 99, 113),
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };

            double[] x = { 2.1, 4.8, 7.5, 10.2, 12.2, 7.5 };
            double[] y = { 5.8, 5.8, 5.8, 5.8, 5.8, 3.2 };
            List<VisioShape> shapes = new();
            for (int i = 0; i < selected.Count && i < x.Length; i++) {
                VisioStencilShape stencil = selected[i];
                string id = "ext-" + i.ToString(System.Globalization.CultureInfo.InvariantCulture);
                VisioShape shape = page.AddStencilShape(stencil, id, x[i], y[i], string.Empty);
                shapes.Add(shape);

                VisioShape label = page.AddTextBox(
                    id + "-label",
                    x[i],
                    y[i] - Math.Max(0.7, shape.Height / 2D + 0.28),
                    1.7,
                    0.42,
                    stencil.Name);
                label.TextStyle = new VisioTextStyle {
                    FontFamily = "Aptos",
                    Size = 11,
                    Color = Color.FromRgb(33, 37, 41),
                    HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                    VerticalAlignment = VisioTextVerticalAlignment.Middle
                };
            }

            for (int i = 0; i < shapes.Count - 1 && i < 4; i++) {
                VisioConnector connector = page.AddConnector(shapes[i], shapes[i + 1], ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
                connector.LineColor = Color.FromRgb(0, 120, 212);
                connector.LineWeight = 0.025;
                connector.EndArrow = EndArrow.Triangle;
            }

            if (shapes.Count > 5) {
                VisioConnector connector = page.AddConnector(shapes[2], shapes[5], ConnectorKind.Straight, VisioSide.Bottom, VisioSide.Top);
                connector.LineColor = Color.FromRgb(111, 66, 193);
                connector.LinePattern = 2;
                connector.LineWeight = 0.025;
                connector.EndArrow = EndArrow.Triangle;
            }

            document.Save();
            IReadOnlyList<string> issues = VisioValidator.Validate(filePath);
            if (issues.Count > 0) {
                throw new InvalidOperationException("Generated external stencil example failed package validation:" + Environment.NewLine + string.Join(Environment.NewLine, issues));
            }

            Console.WriteLine("    Source pack: " + stencilPackPath);
            Console.WriteLine("    Imported masters: " + string.Join(", ", selected.Select(shape => shape.MasterNameU)));
            Console.WriteLine("    Output: " + filePath);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }

        private static IReadOnlyList<VisioStencilShape> SelectStencils(VisioStencilCatalog catalog, params string[] preferredNames) {
            List<VisioStencilShape> selected = new();
            foreach (string preferredName in preferredNames) {
                if (catalog.TryGet(preferredName, out VisioStencilShape? stencil) && stencil != null &&
                    !selected.Any(existing => string.Equals(existing.MasterNameU, stencil.MasterNameU, StringComparison.OrdinalIgnoreCase))) {
                    selected.Add(stencil);
                }
            }

            return selected;
        }
    }
}
