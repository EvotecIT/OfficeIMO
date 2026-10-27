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

            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
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
            PowerPointLayoutBox leftColumn = new(OfficeIMO.OfficeLength.Emus(columns[0].Left.ToEmus()),OfficeIMO.OfficeLength.Emus(bodyTop),OfficeIMO.OfficeLength.Emus(columns[0].Width.ToEmus()),OfficeIMO.OfficeLength.Emus(bodyHeight));
            PowerPointLayoutBox rightColumn = new(OfficeIMO.OfficeLength.Emus(columns[1].Left.ToEmus()),OfficeIMO.OfficeLength.Emus(bodyTop),OfficeIMO.OfficeLength.Emus(columns[1].Width.ToEmus()),OfficeIMO.OfficeLength.Emus(bodyHeight));

            slide.AddTitleCm("Update Picture", marginCm, marginCm, content.Width.ToCentimeters(), 1.3);

            if (File.Exists(imagePath)) {
                PowerPointPicture original = slide.AddImage(imagePath, leftColumn.Left.ToEmus(), leftColumn.Top.ToEmus(), leftColumn.Width.ToEmus(), leftColumn.Height.ToEmus());
                original.FitToBox(1200, 800, crop: false);
                slide.AddTextBox("Original", leftColumn.Left.ToEmus(), leftColumn.Top.ToEmus() + leftColumn.Height.ToEmus() - PowerPointUnits.FromCentimeters(0.8),
                    leftColumn.Width.ToEmus(), PowerPointUnits.FromCentimeters(0.6));

                PowerPointPicture updated = slide.AddImage(imagePath, rightColumn.Left.ToEmus(), rightColumn.Top.ToEmus(), rightColumn.Width.ToEmus(), rightColumn.Height.ToEmus());
                if (File.Exists(newImagePath)) {
                    updated.UpdateImage(newImagePath);
                }
                updated.FitToBox(800, 800, crop: false);
                slide.AddTextBox("Updated via UpdateImage()", rightColumn.Left.ToEmus(),
                    rightColumn.Top.ToEmus() + rightColumn.Height.ToEmus() - PowerPointUnits.FromCentimeters(0.8), rightColumn.Width.ToEmus(), PowerPointUnits.FromCentimeters(0.6));
            } else {
                slide.AddTextBox("Image assets not found.", content.Left.ToEmus(), bodyTop, content.Width.ToEmus(), PowerPointUnits.FromCentimeters(1));
            }

            presentation.Save();
            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }
    }
}
