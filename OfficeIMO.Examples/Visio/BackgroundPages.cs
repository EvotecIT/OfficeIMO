using System;
using System.IO;
using OfficeIMO.Visio;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class BackgroundPages {
        public static void Example_BackgroundPages(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Background pages");
            string filePath = Path.Combine(folderPath, "Background Pages.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage background = document.AddBackgroundPage("Brand background", OfficeLength.Inches(11), OfficeLength.Inches(8.5));
            background.AddRectangle(OfficeLength.Inches(5.5), OfficeLength.Inches(8.05), OfficeLength.Inches(10.5), OfficeLength.Inches(0.45), "OfficeIMO generated diagram")
                .Protect(protection => protection.Size().Position().Text().Selection())
                .FillColor = Color.LightBlue;
            background.AddRectangle(OfficeLength.Inches(5.5), OfficeLength.Inches(0.35), OfficeLength.Inches(10.5), OfficeLength.Inches(0.25), string.Empty).FillColor = Color.LightGray;

            VisioPage architecture = document.AddPage("Architecture", OfficeLength.Inches(11), OfficeLength.Inches(8.5));
            architecture.SetBackgroundPage(background);
            VisioShape api = architecture.AddRectangle(OfficeLength.Inches(3.5), OfficeLength.Inches(4.8), OfficeLength.Inches(2.2), OfficeLength.Inches(1), "API");
            VisioShape worker = architecture.AddRectangle(OfficeLength.Inches(7.5), OfficeLength.Inches(4.8), OfficeLength.Inches(2.2), OfficeLength.Inches(1), "Worker");
            architecture.AddConnector(api, worker, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left).Label = "queue";

            VisioPage operations = document.AddPage("Operations", OfficeLength.Inches(11), OfficeLength.Inches(8.5));
            operations.SetBackgroundPage(background);
            operations.AddRectangle(OfficeLength.Inches(5.5), OfficeLength.Inches(4.8), OfficeLength.Inches(2.2), OfficeLength.Inches(1), "Runbook");

            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
