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
            PowerPointTextBox textBox = slide.AddTextBox("Shapes and images", left: 914400, top: 457200, width: 7315200, height: 914400);
            textBox.FontSize = 28;
            textBox.Color = "1F4E79";

            PowerPointAutoShape card = slide.AddRectangle(914400, 1828800, 3657600, 1828800, "Hero Card")
                .Fill("E7F7FF")
                .Stroke("007ACC", 2);
            card.FillTransparency = 8;
            card.Rotation = 3;

            PowerPointAutoShape accent = slide.AddEllipse(3657600, 1828800, 1828800, 1828800, "Accent")
                .Fill("FDEBD0")
                .Stroke("D35400", 1.5);
            accent.HorizontalFlip = true;
            accent.SendToBack();

            PowerPointAutoShape connector = slide.AddLine(914400, 3886200, 5486400, 3886200, "Connector");
            connector.Stroke("404040", 2);

            PowerPointPicture picture = slide.AddPicture(imagePath, left: 6400800, top: 1828800, width: 2286000, height: 2286000);
            picture.FitToBox(1200, 1200, crop: true);

            PowerPointShape? shape = slide.GetShape("Hero Card");
            Console.WriteLine("Found shape: " + shape?.Name);
            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}

