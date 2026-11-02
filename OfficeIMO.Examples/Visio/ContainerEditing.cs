using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class ContainerEditing {
        public static void Example_ContainerEditing(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Container editing");
            string filePath = Path.Combine(folderPath, "Container Editing.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Application Containers", OfficeLength.Inches(11), OfficeLength.Inches(8.5));

            VisioShape api = page.AddStencilShape(VisioStencils.Network.Get("server"), "api", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5.5, page.DefaultUnit)), "API");
            VisioShape worker = page.AddStencilShape(VisioStencils.Network.Get("server"), "worker", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5.5, page.DefaultUnit)), "Worker");
            VisioShape queue = page.AddStencilShape(VisioStencils.Flowchart.Get("data"), "queue", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3.6, page.DefaultUnit)), "Queue");

            page.AddConnector(api, queue, ConnectorKind.Dynamic, VisioSide.Bottom, VisioSide.Top).Label = "publish";
            page.AddConnector(queue, worker, ConnectorKind.Dynamic, VisioSide.Top, VisioSide.Bottom).Label = "consume";

            VisioShape container = page.AddContainer("app-tier", "Application tier", new[] { api, worker, queue }, new VisioContainerOptions {
                Margin = 0.35,
                HeadingHeight = 0.4,
                FillColor = Color.LightCyan,
                LineColor = Color.DodgerBlue
            });
            container.SetUserCell("OfficeIMO.Role", "Tier", "STR", prompt: "Semantic role");

            page.SelectContainers().Stroke(Color.DodgerBlue, 0.02);
            page.SelectWithUserCell("OfficeIMO.Role", "Tier").UserCell("OfficeIMO.Reviewed", "Yes", "STR");

            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
