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
            background.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5.5, background.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(8.05, background.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(10.5, background.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.45, background.DefaultUnit)), "OfficeIMO generated diagram")
                .Protect(protection => protection.Size().Position().Text().Selection())
                .FillColor = Color.LightBlue;
            background.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5.5, background.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.35, background.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(10.5, background.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(0.25, background.DefaultUnit)), string.Empty).FillColor = Color.LightGray;

            VisioPage architecture = document.AddPage("Architecture", OfficeLength.Inches(11), OfficeLength.Inches(8.5));
            architecture.SetBackgroundPage(background);
            VisioShape api = architecture.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3.5, architecture.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4.8, architecture.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.2, architecture.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, architecture.DefaultUnit)), "API");
            VisioShape worker = architecture.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7.5, architecture.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4.8, architecture.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.2, architecture.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, architecture.DefaultUnit)), "Worker");
            architecture.AddConnector(api, worker, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left).Label = "queue";

            VisioPage operations = document.AddPage("Operations", OfficeLength.Inches(11), OfficeLength.Inches(8.5));
            operations.SetBackgroundPage(background);
            operations.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5.5, operations.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4.8, operations.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.2, operations.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, operations.DefaultUnit)), "Runbook");

            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
