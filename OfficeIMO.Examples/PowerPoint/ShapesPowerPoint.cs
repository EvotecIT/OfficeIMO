using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates retrieving, positioning, and removing shapes.
    /// </summary>
    public static class ShapesPowerPoint {
        public static void Example_PowerPointShapes(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Shape operations");
            string filePath = Path.Combine(folderPath, "Shape Operations.pptx");
            string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "Images", "BackgroundImage.png");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            PowerPointSlide slide = presentation.AddSlide();
            PPTextBox textBox = slide.AddTextBox("Hello", left: 914400, top: 914400, width: 1828800, height: 914400);
            PPPicture picture = slide.AddPicture(imagePath, left: 2743200, top: 914400, width: 1828800, height: 1828800);

            // Move the textbox 1 inch to the right
            textBox.Left += 914400;

            PPShape? shape = slide.GetShape("TextBox 1");
            Console.WriteLine("Found shape: " + shape?.Name);
            slide.RemoveShape(picture);
            presentation.Save();
        }
    }
}
