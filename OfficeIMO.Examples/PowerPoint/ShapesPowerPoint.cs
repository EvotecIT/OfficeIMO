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
            const double marginCm = 1.5;
            const double gutterCm = 1.0;
            const double titleHeightCm = 1.4;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            double slideHeightCm = presentation.SlideSize.HeightCm;
            double bodyTopCm = 3.4;
            double bodyHeightCm = slideHeightCm - bodyTopCm - marginCm;
            PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumnsCm(2, marginCm, gutterCm);
            PowerPointLayoutBox leftColumn = PowerPointLayoutBox.FromCentimeters(columns[0].LeftCm, bodyTopCm, columns[0].WidthCm, bodyHeightCm);
            PowerPointLayoutBox rightColumn = PowerPointLayoutBox.FromCentimeters(columns[1].LeftCm, bodyTopCm, columns[1].WidthCm, bodyHeightCm);

            if (File.Exists(imagePath)) {
                slide.SetBackgroundImage(imagePath);
            }

            PowerPointTextBox textBox = slide.AddTitleCm("Shapes and images", marginCm, marginCm, content.WidthCm, titleHeightCm);
            textBox.FontSize = 30;
            textBox.Color = "1F4E79";

            const double gapCm = 0.4;
            double cardHeightCm = (leftColumn.HeightCm - gapCm) / 2;

            PowerPointAutoShape card = slide.AddRectangleCm(leftColumn.LeftCm, leftColumn.TopCm, leftColumn.WidthCm, cardHeightCm, "Hero Card")
                .Fill("E7F7FF")
                .Stroke("007ACC", 2);
            card.FillTransparency = 6;

            slide.AddTextBoxCm("Card with accent and baseline", leftColumn.LeftCm + 0.5,
                leftColumn.TopCm + 0.4, leftColumn.WidthCm - 1.0, 1.0);

            const double badgeSizeCm = 2.2;
            slide.AddEllipseCm(
                    leftColumn.LeftCm + leftColumn.WidthCm - badgeSizeCm - 0.6,
                    leftColumn.TopCm + 0.6,
                    badgeSizeCm,
                    badgeSizeCm,
                    "Badge")
                .Fill("FDEBD0")
                .Stroke("D35400", 1.5);

            double baselineTopCm = leftColumn.TopCm + cardHeightCm + gapCm + 0.4;
            PowerPointAutoShape connector = slide.AddLineCm(leftColumn.LeftCm, baselineTopCm, leftColumn.LeftCm + leftColumn.WidthCm, baselineTopCm,
                "Connector");
            connector.Stroke("404040", 1.5);

            PowerPointTextBox label = slide.AddTextBoxCm("Shapes stay aligned within layout columns.", leftColumn.LeftCm,
                baselineTopCm + 0.3, leftColumn.WidthCm, 2.0);
            label.FontSize = 16;

            if (File.Exists(imagePath)) {
                PowerPointPicture picture = slide.AddPictureCm(imagePath, rightColumn.LeftCm, rightColumn.TopCm, rightColumn.WidthCm,
                    rightColumn.HeightCm);
                picture.FitToBox(1200, 800, crop: true);
                PowerPointTextBox caption = slide.AddTextBoxCm("Image scaled to fill column", rightColumn.LeftCm,
                    rightColumn.TopCm + rightColumn.HeightCm - 0.7, rightColumn.WidthCm, 0.6);
                caption.FontSize = 12;
                caption.Color = "666666";
            } else {
                slide.AddTextBoxCm("(image placeholder)", rightColumn.LeftCm, rightColumn.TopCm, rightColumn.WidthCm, 2.5);
            }

            PowerPointShape? shape = slide.GetShape("Hero Card");
            Console.WriteLine("Found shape: " + shape?.Name);
            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
