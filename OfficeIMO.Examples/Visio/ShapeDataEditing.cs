using System;
using System.IO;
using OfficeIMO.Visio;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class ShapeDataEditing {
        public static void Example_ShapeDataEditing(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Shape Data editing");
            string filePath = Path.Combine(folderPath, "Shape Data Editing.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Shape Data", OfficeLength.Inches(8.5), OfficeLength.Inches(6));

            VisioShape api = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "API");
            api.SetShapeData("Owner", "Platform", "Owner", VisioShapeDataType.String, "Owning support team");
            api.SetShapeData("MonthlyCost", "1250", "Monthly cost", VisioShapeDataType.Currency, "Estimated monthly cost", "$#,##0");

            VisioShape database = page.AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2.2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Database");
            database.SetShapeData("Owner", "Data", "Owner", VisioShapeDataType.String, "Owning support team");
            database.SetShapeData("MonthlyCost", "2200", "Monthly cost", VisioShapeDataType.Currency, "Estimated monthly cost", "$#,##0");

            page.SelectWithShapeData("Owner", "Platform")
                .Fill(Color.LightBlue)
                .ShapeData("Reviewed", "Yes", "Reviewed", VisioShapeDataType.Boolean, "Architecture review complete");

            document.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
