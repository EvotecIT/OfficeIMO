using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates slide removal and reordering.
    /// </summary>
    public static class SlidesManagementPowerPoint {
        public static void Example_SlidesManagement(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Slides management");
            string filePath = Path.Combine(folderPath, "Slides Management.pptx");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            presentation.AddSlide().AddTextBox("Slide 1");
            presentation.AddSlide().AddTextBox("Slide 2");
            presentation.AddSlide().AddTextBox("Slide 3");

            presentation.MoveSlide(2, 0);
            presentation.RemoveSlide(1);
            presentation.Save();
        }
    }
}
