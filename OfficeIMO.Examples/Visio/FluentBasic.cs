using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;

namespace OfficeIMO.Examples.Visio {
    public static class FluentBasicVisio {
        public static void Example_FluentBasicVisio(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Fluent basic diagram");
            string filePath = Path.Combine(folderPath, "Fluent Basic Visio.vsdx");

            var doc = VisioDocument.Create(filePath);
            doc.AsFluent()
               .Info(i => i.Title("Fluent Visio").Author("OfficeIMO"))
               .AddPage("Page-1", p => p
                   .AddRectangle("S1", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Start")
                   .AddDiamond("D1", OfficeLength.Inches(4), OfficeLength.Inches(1.5), OfficeLength.Inches(2), OfficeLength.Inches(2), "Decision")
                   .AddEllipse("E1", OfficeLength.Inches(7), OfficeLength.Inches(1.5), OfficeLength.Inches(2), OfficeLength.Inches(1), "End")
                   .AddConnector("S1", "D1", configure: c => c.RightAngle().ArrowEnd(EndArrow.Triangle))
                   .AddConnector("D1", "E1", configure: c => c.RightAngle().ArrowEnd(EndArrow.Triangle).Label("Yes"))
               )
               .End();
            doc.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

