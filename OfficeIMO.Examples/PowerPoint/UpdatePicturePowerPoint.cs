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
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Centimeters(marginCm));
            double slideHeightCm = presentation.SlideSize.Height.ToCentimeters();
            double bodyTopCm = 3.4;
            double bodyHeightCm = slideHeightCm - bodyTopCm - marginCm;
            long bodyTop = PowerPointUnits.FromCentimeters(bodyTopCm);
            long bodyHeight = PowerPointUnits.FromCentimeters(bodyHeightCm);
            PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumns(2,OfficeIMO.OfficeLength.Centimeters(marginCm),OfficeIMO.OfficeLength.Centimeters(gutterCm));
            PowerPointLayoutBox leftColumn = new(OfficeIMO.OfficeLength.Emus(columns[0].Left.ToEmus()),OfficeIMO.OfficeLength.Emus(bodyTop),OfficeIMO.OfficeLength.Emus(columns[0].Width.ToEmus()),OfficeIMO.OfficeLength.Emus(bodyHeight));
            PowerPointLayoutBox rightColumn = new(OfficeIMO.OfficeLength.Emus(columns[1].Left.ToEmus()),OfficeIMO.OfficeLength.Emus(bodyTop),OfficeIMO.OfficeLength.Emus(columns[1].Width.ToEmus()),OfficeIMO.OfficeLength.Emus(bodyHeight));

            slide.AddTitle("Update Picture",global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.3));

            if (File.Exists(imagePath)) {
                PowerPointPicture original = slide.AddImage(imagePath, leftColumn.Left, leftColumn.Top, leftColumn.Width, leftColumn.Height);
                original.FitToBox(1200, 800, crop: false);
                slide.AddTextBox("Original", leftColumn.Left, global::OfficeIMO.OfficeLength.Emus(leftColumn.Top.ToEmus() + leftColumn.Height.ToEmus() - PowerPointUnits.FromCentimeters(0.8)),
                    leftColumn.Width, global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromCentimeters(0.6)));

                PowerPointPicture updated = slide.AddImage(imagePath, rightColumn.Left, rightColumn.Top, rightColumn.Width, rightColumn.Height);
                if (File.Exists(newImagePath)) {
                    updated.UpdateImage(newImagePath);
                }
                updated.FitToBox(800, 800, crop: false);
                slide.AddTextBox("Updated via UpdateImage()", rightColumn.Left,
                    global::OfficeIMO.OfficeLength.Emus(rightColumn.Top.ToEmus() + rightColumn.Height.ToEmus() - PowerPointUnits.FromCentimeters(0.8)), rightColumn.Width, global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromCentimeters(0.6)));
            } else {
                slide.AddTextBox("Image assets not found.", content.Left, global::OfficeIMO.OfficeLength.Emus(bodyTop), content.Width, global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromCentimeters(1)));
            }

            presentation.Save();
            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }
    }
}
