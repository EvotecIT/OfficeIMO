using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates basic <see cref="PowerPointDocument"/> usage.
    /// </summary>
    public static class BasicPowerPointDocument {
        public static void Example_BasicPowerPoint(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Creating basic presentation");
            string filePath = Path.Combine(folderPath, "Basic PowerPoint.pptx");

            PowerPointDocument document = new();
            PowerPointSlide slide = document.AddSlide("Slide1");
            slide.Shapes.Add(new PowerPointShape("Shape1"));

            // Saving not yet implemented; placeholder for future development.
        }
    }
}
