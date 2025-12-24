using System;
using System.IO;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Fluent;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates fluent API for building presentations.
    /// </summary>
    public static class FluentPowerPoint {
        public static void Example_FluentPowerPoint(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Creating presentation with fluent API");
            string filePath = Path.Combine(folderPath, "FluentPowerPoint.pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AsFluent()
                    .Slide(0, 0, s => {
                        s.Title("Fluent Presentation", tb => {
                            tb.FontSize = 32;
                            tb.Color = "1F4E79";
                        });
                        s.TextBox("Hello from fluent API", tb => {
                            tb.AddBullet("Built with builders");
                            tb.AddBullet("Configurable content");
                        });
                        s.Numbered("Step one", "Step two");
                        s.Shape(A.ShapeTypeValues.Rectangle, 914400, 3657600, 4572000, 1143000,
                            shape => shape.Fill("E7F7FF").Stroke("007ACC", 2));
                        s.Notes("Example notes");
                    })
                    .Slide(s => s.Title("Second Slide"))
                    .End()
                    .Save();
            }

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
