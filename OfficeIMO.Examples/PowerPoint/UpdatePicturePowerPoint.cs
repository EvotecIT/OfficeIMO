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
            const double marginCm = 1.5;
            const double gutterCm = 1.0;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            double slideHeightCm = presentation.SlideSize.HeightCm;
            double bodyTopCm = 3.4;
            double bodyHeightCm = slideHeightCm - bodyTopCm - marginCm;
            long bodyTop = PowerPointUnits.FromCentimeters(bodyTopCm);
            long bodyHeight = PowerPointUnits.FromCentimeters(bodyHeightCm);
            PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumnsCm(2, marginCm, gutterCm);
            PowerPointLayoutBox leftColumn = new(columns[0].Left, bodyTop, columns[0].Width, bodyHeight);
            PowerPointLayoutBox rightColumn = new(columns[1].Left, bodyTop, columns[1].Width, bodyHeight);

            slide.AddTitleCm("Update Picture", marginCm, marginCm, content.WidthCm, 1.3);

            if (File.Exists(imagePath)) {
                PowerPointPicture original = slide.AddPicture(imagePath, leftColumn.Left, leftColumn.Top, leftColumn.Width, leftColumn.Height);
                original.FitToBox(1200, 800, crop: false);
                slide.AddTextBox("Original", leftColumn.Left, leftColumn.Top + leftColumn.Height - PowerPointUnits.Cm(0.8),
                    leftColumn.Width, PowerPointUnits.Cm(0.6));

                PowerPointPicture updated = slide.AddPicture(imagePath, rightColumn.Left, rightColumn.Top, rightColumn.Width, rightColumn.Height);
                if (File.Exists(newImagePath)) {
                    updated.UpdateImage(newImagePath);
                }
                updated.FitToBox(800, 800, crop: false);
                slide.AddTextBox("Updated via UpdateImage()", rightColumn.Left,
                    rightColumn.Top + rightColumn.Height - PowerPointUnits.Cm(0.8), rightColumn.Width, PowerPointUnits.Cm(0.6));
            } else {
                slide.AddTextBox("Image assets not found.", content.Left, bodyTop, content.Width, PowerPointUnits.Cm(1));
            }

            presentation.Save();
            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
