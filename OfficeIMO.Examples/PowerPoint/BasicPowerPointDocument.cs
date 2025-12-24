using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates basic <see cref="PowerPointPresentation"/> usage.
    /// </summary>
    public static class BasicPowerPointDocument {
        public static void Example_BasicPowerPoint(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Creating basic presentation");
            string filePath = Path.Combine(folderPath, "Basic PowerPoint.pptx");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            Console.WriteLine($"Theme: {presentation.ThemeName}");
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox text = slide.AddTextBox("Hello World");
            text.AddBullet("Bullet 1");
            //slide.Notes.Text = "Example notes";
            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}

