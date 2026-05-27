using System;
using System.IO;
using OfficeIMO.Visio;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Examples.Visio {
    public static class ShapeDataEditing {
        public static void Example_ShapeDataEditing(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Shape Data editing");
            string filePath = Path.Combine(folderPath, "Shape Data Editing.vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Shape Data", 8.5, 6);

            VisioShape api = page.AddRectangle(2.5, 4, 2.2, 1, "API");
            api.SetShapeData("Owner", "Platform", "Owner", VisioShapeDataType.String, "Owning support team");
            api.SetShapeData("MonthlyCost", "1250", "Monthly cost", VisioShapeDataType.Currency, "Estimated monthly cost", "$#,##0");

            VisioShape database = page.AddRectangle(6, 4, 2.2, 1, "Database");
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
