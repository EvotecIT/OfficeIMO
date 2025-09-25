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
               .Page("Page-1", p => p
                   .Rect("S1", 1, 1, 2, 1, "Start")
                   .Diamond("D1", 4, 1.5, 2, 2, "Decision")
                   .Ellipse("E1", 7, 1.5, 2, 1, "End")
                   .Connect("S1", "D1", c => c.RightAngle().ArrowEnd(EndArrow.Triangle))
                   .Connect("D1", "E1", c => c.RightAngle().ArrowEnd(EndArrow.Triangle).Label("Yes"))
               )
               .End();
            doc.Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

