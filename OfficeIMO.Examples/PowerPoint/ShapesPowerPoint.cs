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
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            double slideHeightCm = presentation.SlideSize.HeightCm;
            double bodyTopCm = 3.4;
            double bodyHeightCm = slideHeightCm - bodyTopCm - marginCm;
            long bodyTop = PowerPointUnits.FromCentimeters(bodyTopCm);
            long bodyHeight = PowerPointUnits.FromCentimeters(bodyHeightCm);
            PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumnsCm(2, marginCm, gutterCm);
            PowerPointLayoutBox leftColumn = new(columns[0].Left, bodyTop, columns[0].Width, bodyHeight);
            PowerPointLayoutBox rightColumn = new(columns[1].Left, bodyTop, columns[1].Width, bodyHeight);

            PowerPointTextBox textBox = slide.AddTitleCm("Shapes and images", marginCm, marginCm, content.WidthCm, 1.3);
            textBox.FontSize = 30;
            textBox.Color = "1F4E79";

            long gap = PowerPointUnits.Cm(0.4);
            long cardHeight = (leftColumn.Height - gap) / 2;

            PowerPointAutoShape card = slide.AddRectangle(leftColumn.Left, leftColumn.Top, leftColumn.Width, cardHeight, "Hero Card")
                .Fill("E7F7FF")
                .Stroke("007ACC", 2);
            card.FillTransparency = 6;

            slide.AddTextBox("Card with accent and baseline", leftColumn.Left + PowerPointUnits.Cm(0.5),
                leftColumn.Top + PowerPointUnits.Cm(0.4), leftColumn.Width - PowerPointUnits.Cm(1), PowerPointUnits.Cm(1));

            long badgeSize = PowerPointUnits.Cm(2.2);
            slide.AddEllipse(
                    leftColumn.Left + leftColumn.Width - badgeSize - PowerPointUnits.Cm(0.6),
                    leftColumn.Top + PowerPointUnits.Cm(0.6),
                    badgeSize,
                    badgeSize,
                    "Badge")
                .Fill("FDEBD0")
                .Stroke("D35400", 1.5);

            long baselineTop = leftColumn.Top + cardHeight + gap + PowerPointUnits.Cm(0.4);
            PowerPointAutoShape connector = slide.AddLine(leftColumn.Left, baselineTop, leftColumn.Left + leftColumn.Width, baselineTop, "Connector");
            connector.Stroke("404040", 1.5);

            PowerPointTextBox label = slide.AddTextBox("Shapes stay aligned within layout columns.",
                leftColumn.Left, baselineTop + PowerPointUnits.Cm(0.3), leftColumn.Width, PowerPointUnits.Cm(2));
            label.FontSize = 16;

            if (File.Exists(imagePath)) {
                PowerPointPicture picture = slide.AddPicture(imagePath, rightColumn.Left, rightColumn.Top, rightColumn.Width, rightColumn.Height);
                picture.FitToBox(1200, 800, crop: false);
                PowerPointTextBox caption = slide.AddTextBox("Image scaled to fit column", rightColumn.Left,
                    rightColumn.Top + rightColumn.Height - PowerPointUnits.Cm(0.9), rightColumn.Width, PowerPointUnits.Cm(0.6));
                caption.FontSize = 12;
                caption.Color = "666666";
            }

            PowerPointShape? shape = slide.GetShape("Hero Card");
            Console.WriteLine("Found shape: " + shape?.Name);
            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
