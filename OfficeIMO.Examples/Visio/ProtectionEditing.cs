using System;
using System.IO;
using OfficeIMO.Visio;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class ProtectionEditing {
        public static void Example_ProtectionEditing(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Protection editing");
            string filePath = Path.Combine(folderPath, "Protection Editing.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Protected Diagram", OfficeLength.Inches(11), OfficeLength.Inches(8.5));

            VisioShape background = page.AddRectangle(OfficeLength.Inches(5.5), OfficeLength.Inches(4.25), OfficeLength.Inches(9.5), OfficeLength.Inches(6.5), "Generated architecture zone");
            background.FillColor = Color.LightCyan;
            background.LineColor = Color.SteelBlue;
            background.Protect(protection => protection.Size().Position().Selection().Formatting());

            VisioShape api = page.AddRectangle(OfficeLength.Inches(3.5), OfficeLength.Inches(5.1), OfficeLength.Inches(2.2), OfficeLength.Inches(1), "API");
            VisioShape worker = page.AddRectangle(OfficeLength.Inches(7.5), OfficeLength.Inches(5.1), OfficeLength.Inches(2.2), OfficeLength.Inches(1), "Worker");
            VisioShape database = page.AddRectangle(OfficeLength.Inches(5.5), OfficeLength.Inches(2.8), OfficeLength.Inches(2.2), OfficeLength.Inches(1), "Database");

            page.AddConnector(api, worker, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left).Label = "queue";
            page.AddConnector(worker, database, ConnectorKind.Dynamic, VisioSide.Bottom, VisioSide.Top).Label = "write";
            page.AddConnector(api, database, ConnectorKind.Dynamic, VisioSide.Bottom, VisioSide.Top).Label = "read";

            page.SelectContainingText("API").ShapeData("Owner", "Platform");
            page.SelectWithData("Owner", "Platform")
                .Protect(protection => protection.Text().Deletion())
                .Fill(Color.LightYellow);
            page.SelectConnectors(connector => connector.Label != null)
                .LockEndpoints()
                .Protect(protection => protection.Text().Deletion());

            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
