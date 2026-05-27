using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class StencilCatalog {
        public static void Example_StencilCatalog(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Native stencil catalog");
            string filePath = Path.Combine(folderPath, "Native Stencil Catalog.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Stencil Catalog", 29.7, 21.0, VisioMeasurementUnit.Centimeters);

            VisioShape start = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "start", 5, 15, "Receive request");
            VisioShape decision = page.AddStencilShape(VisioStencils.Flowchart, "branch", "decision", 13, 15, "Approved?");
            VisioStencilShape archiveStencil = VisioStencils.All.Search("data-store").First();
            VisioShape storage = page.AddStencilShape(archiveStencil, "storage", 21, 15, "Archive");
            VisioStencilShape wireless = VisioStencils.Network.Search("access-point").First();
            VisioShape accessPoint = page.AddStencilShape(wireless, "wifi", 21, 9, "Wi-Fi");
            VisioStencilCatalog custom = VisioStencilCatalog.Create("Custom Infrastructure", catalog => catalog
                .Add("custom.cache", "Cache", "Process", "Infrastructure", 2.2, 1.0, "redis", "memory-store"));
            string manifestPath = Path.Combine(folderPath, "Custom Infrastructure.officeimo-visio-stencils.xml");
            custom.Save(manifestPath);
            VisioStencilCatalog reusableCustom = VisioStencilCatalog.Load(manifestPath);
            VisioShape cache = page.AddStencilShape(custom, "redis", "cache", 13, 9, "Cache");

            foreach (VisioShape shape in new[] { start, decision, storage, accessPoint, cache }) {
                shape.FillColor = Color.FromRgb(231, 244, 255);
                shape.LineColor = Color.FromRgb(0, 120, 212);
                shape.LineWeight = 0.02;
            }

            page.AddConnector(start, decision, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left).EndArrow = EndArrow.Triangle;
            page.AddConnector(decision, storage, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left).EndArrow = EndArrow.Triangle;

            Console.WriteLine("    Categories: " + string.Join(", ", VisioStencils.All.Categories));
            Console.WriteLine("    Network search 'access-point': " + wireless.Name + " (" + wireless.IconNameU + ")");
            Console.WriteLine("    Custom manifest: " + manifestPath);
            Console.WriteLine("    Custom search 'redis': " + reusableCustom.Search("redis").First().Name);

            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
