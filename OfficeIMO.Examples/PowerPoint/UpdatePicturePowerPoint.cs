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
            string newImagePath = Path.Combine(Directory.GetCurrentDirectory(), "Images", "EvotecLogo.png");
            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointPicture picture = slide.AddPicture(imagePath, left: 914400, top: 914400, width: 7315200, height: 3657600);
            Console.WriteLine("Original type: " + picture.ContentType);

            picture.FitToBox(1200, 800, crop: true);
            picture.Crop(5, 5, 5, 5);
            picture.UpdateImage(newImagePath);
            picture.ResetCrop();
            picture.FitToBox(800, 800, crop: false);
            Console.WriteLine("Updated type: " + picture.ContentType);

            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
