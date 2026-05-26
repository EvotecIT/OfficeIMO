using System;
using System.IO;
using OfficeIMO.Visio;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class BackgroundPages {
        public static void Example_BackgroundPages(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Background pages");
            string filePath = Path.Combine(folderPath, "Background Pages.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage background = document.AddBackgroundPage("Brand background", 11, 8.5);
            background.AddRectangle(5.5, 8.05, 10.5, 0.45, "OfficeIMO generated diagram")
                .Protect(protection => protection.Size().Position().Text().Selection())
                .FillColor = Color.LightBlue;
            background.AddRectangle(5.5, 0.35, 10.5, 0.25, string.Empty).FillColor = Color.LightGray;

            VisioPage architecture = document.AddPage("Architecture", 11, 8.5);
            architecture.SetBackgroundPage(background);
            VisioShape api = architecture.AddRectangle(3.5, 4.8, 2.2, 1, "API");
            VisioShape worker = architecture.AddRectangle(7.5, 4.8, 2.2, 1, "Worker");
            architecture.AddConnector(api, worker, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left).Label = "queue";

            VisioPage operations = document.AddPage("Operations", 11, 8.5);
            operations.SetBackgroundPage(background);
            operations.AddRectangle(5.5, 4.8, 2.2, 1, "Runbook");

            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
