using System;
using System.IO;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates retrieving, positioning, and removing shapes.
    /// </summary>
    public static class ShapesPowerPoint {
        public static void Example_PowerPointShapes(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Shape operations");
            string filePath = Path.Combine(folderPath, "Shape Operations.pptx");
            string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "Images", "BackgroundImage.png");
            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
            PowerPointSlide slide = presentation.AddSlide();
            const double marginCm = 1.5;
            const double gutterCm = 1.0;
            const double titleHeightCm = 1.4;
            const double bodyGapCm = 0.8;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Centimeters(marginCm));
            PowerPointLayoutBox titleBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(titleHeightCm));
            PowerPointLayoutBox bodyBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),                OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + titleHeightCm + bodyGapCm),                OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),                OfficeIMO.OfficeLength.Centimeters(content.Height.ToCentimeters() - titleHeightCm - bodyGapCm));
            PowerPointLayoutBox[] columns = bodyBox.SplitColumns(2,OfficeIMO.OfficeLength.Centimeters(gutterCm));
            PowerPointLayoutBox leftColumn = columns[0];
            PowerPointLayoutBox rightColumn = columns[1];

            if (File.Exists(imagePath)) {
                slide.SetBackgroundImage(imagePath);
            }

            PowerPointTextBox textBox = slide.AddTitle("Shapes and images",titleBox);
            textBox.FontSize = global::OfficeIMO.OfficeLength.Points(30);
            textBox.Color = "1F4E79";

            const double captionGapCm = 0.6;
            const double captionHeightCm = 1.2;
            double canvasHeightCm = leftColumn.Height.ToCentimeters() - captionGapCm - captionHeightCm;
            PowerPointLayoutBox canvasBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(leftColumn.Left.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(leftColumn.Top.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(leftColumn.Width.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(canvasHeightCm));
            PowerPointLayoutBox captionBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(leftColumn.Left.ToCentimeters()),                OfficeIMO.OfficeLength.Centimeters(leftColumn.Top.ToCentimeters() + canvasHeightCm + captionGapCm),                OfficeIMO.OfficeLength.Centimeters(leftColumn.Width.ToCentimeters()),                OfficeIMO.OfficeLength.Centimeters(captionHeightCm));

            PowerPointAutoShape card = slide.AddRectangle(
                    canvasBox.Left, canvasBox.Top, canvasBox.Width, OfficeIMO.OfficeLength.Centimeters(canvasBox.Height.ToCentimeters() * 0.6),
                    "Hero Card")
                .Fill("E7F7FF")
                .Stroke("007ACC", global::OfficeIMO.OfficeLength.Points(2));
            card.FillTransparency = 6;
            card.Rotation = -1.5;

            PowerPointAutoShape accent = slide.AddRectangle(
                    OfficeIMO.OfficeLength.Centimeters(canvasBox.Left.ToCentimeters() + 0.7),
                    OfficeIMO.OfficeLength.Centimeters(canvasBox.Top.ToCentimeters() + 0.7),
                    OfficeIMO.OfficeLength.Centimeters(canvasBox.Width.ToCentimeters() * 0.55),
                    OfficeIMO.OfficeLength.Centimeters(canvasBox.Height.ToCentimeters() * 0.35),
                    "Accent Panel")
                .Fill("FFF4E5")
                .Stroke("C48A00", global::OfficeIMO.OfficeLength.Points(1.5));
            accent.Rotation = 2;
            accent.BringToFront();

            const double badgeSizeCm = 2.1;
            slide.AddEllipse(
                    OfficeIMO.OfficeLength.Centimeters(canvasBox.Left.ToCentimeters() + canvasBox.Width.ToCentimeters() - badgeSizeCm - 0.4),
                    OfficeIMO.OfficeLength.Centimeters(canvasBox.Top.ToCentimeters() + 0.4),
                    OfficeIMO.OfficeLength.Centimeters(badgeSizeCm),
                    OfficeIMO.OfficeLength.Centimeters(badgeSizeCm),
                    "Badge")
                .Fill("FDEBD0")
                .Stroke("D35400", global::OfficeIMO.OfficeLength.Points(1.5));

            PowerPointAutoShape arrow = slide.AddShape(
                    OfficePresetShapeType.RightArrow,
                    OfficeIMO.OfficeLength.Centimeters(canvasBox.Left.ToCentimeters() + 0.4),
                    OfficeIMO.OfficeLength.Centimeters(canvasBox.Top.ToCentimeters() + canvasBox.Height.ToCentimeters() * 0.68),
                    OfficeIMO.OfficeLength.Centimeters(canvasBox.Width.ToCentimeters() * 0.6),
                    OfficeIMO.OfficeLength.Centimeters(canvasBox.Height.ToCentimeters() * 0.25),
                    "Flow Arrow")
                .Fill("D9EAD3")
                .Stroke("6AA84F", global::OfficeIMO.OfficeLength.Points(1.5));
            arrow.Rotation = -1;

            PowerPointTextBox label = slide.AddTextBox(
                "Rotation, layering, and named shapes for lookup.",                captionBox);
            label.FontSize = global::OfficeIMO.OfficeLength.Points(14);
            label.Color = "1F4E79";

            if (File.Exists(imagePath)) {
                PowerPointPicture picture = slide.AddImage(imagePath,rightColumn);
                picture.FitToBox(1200, 800, crop: true);
            } else {
                PowerPointTextBox placeholder = slide.AddTextBox("(image placeholder)",rightColumn);
                placeholder.FillColor = "F3F3F3";
                placeholder.OutlineColor = "CCCCCC";
                placeholder.TextVerticalAlignment = PowerPointTextVerticalAlignment.Center;
            }

            PowerPointShape? shape = slide.GetShape("Hero Card");
            if (shape != null) {
                shape.BringToFront();
            }
            Console.WriteLine("Found shape: " + shape?.Name);
            presentation.Save();

            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }
    }
}
