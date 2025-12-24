using System;
using System.IO;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Fluent;

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
                    .Slide(0, 0)
                        .Title("Fluent Presentation")
                        .TextBox("Hello from fluent API")
                        .Bullets("First", "Second")
                        //.Notes("Example notes")
                        .End()
                    .Slide(s => s.Title("Second Slide"))
                    .End()
                    .Save();
            }

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
