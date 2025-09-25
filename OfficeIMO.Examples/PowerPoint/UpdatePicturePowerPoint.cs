using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates replacing an image in a PowerPoint slide.
    /// </summary>
    public static class UpdatePicturePowerPoint {
        public static void Example_PowerPointUpdatePicture(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Update picture");
            string filePath = Path.Combine(folderPath, "Update Picture.pptx");
            string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "Images", "BackgroundImage.png");
            string newImagePath = Path.Combine(Directory.GetCurrentDirectory(), "Images", "Kulek.jpg");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointPicture picture = slide.AddPicture(imagePath);
            Console.WriteLine("Original type: " + picture.ContentType);

            using FileStream stream = new(newImagePath, FileMode.Open, FileAccess.Read);
            picture.UpdateImage(stream, ImagePartType.Jpeg);
            Console.WriteLine("Updated type: " + picture.ContentType);

            presentation.Save();
        }
    }
}
