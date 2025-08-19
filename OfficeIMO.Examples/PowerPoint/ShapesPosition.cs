using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates positioning and sizing of shapes.
    /// </summary>
    public static class ShapesPosition {
        public static void Example_ShapesPosition(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Shapes position and size");
            string filePath = Path.Combine(folderPath, "Shapes Position.pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTextBox("Text box", left: 1000000L, top: 1000000L, width: 3000000L, height: 1000000L);
            slide.AddPicture(imagePath, left: 3000000L, top: 1000000L, width: 2000000L, height: 2000000L);
            slide.AddTable(2, 2, left: 1000000L, top: 2500000L, width: 4000000L, height: 1000000L);
            presentation.Save();
        }
    }
}
