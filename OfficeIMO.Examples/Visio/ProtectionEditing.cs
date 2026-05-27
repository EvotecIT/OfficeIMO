using System;
using System.IO;
using OfficeIMO.Visio;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class ProtectionEditing {
        public static void Example_ProtectionEditing(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Protection editing");
            string filePath = Path.Combine(folderPath, "Protection Editing.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Protected Diagram", 11, 8.5);

            VisioShape background = page.AddRectangle(5.5, 4.25, 9.5, 6.5, "Generated architecture zone");
            background.FillColor = Color.LightCyan;
            background.LineColor = Color.SteelBlue;
            background.Protect(protection => protection.Size().Position().Selection().Formatting());

            VisioShape api = page.AddRectangle(3.5, 5.1, 2.2, 1, "API");
            VisioShape worker = page.AddRectangle(7.5, 5.1, 2.2, 1, "Worker");
            VisioShape database = page.AddRectangle(5.5, 2.8, 2.2, 1, "Database");

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
