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
            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            PowerPointSlide slide = presentation.AddSlide();
            const double marginCm = 1.5;
            const double gutterCm = 1.0;
            const double titleHeightCm = 1.4;
            const double bodyGapCm = 0.8;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            PowerPointLayoutBox titleBox = PowerPointLayoutBox.FromCentimeters(
                content.LeftCm, content.TopCm, content.WidthCm, titleHeightCm);
            PowerPointLayoutBox bodyBox = PowerPointLayoutBox.FromCentimeters(
                content.LeftCm,
                content.TopCm + titleHeightCm + bodyGapCm,
                content.WidthCm,
                content.HeightCm - titleHeightCm - bodyGapCm);
            PowerPointLayoutBox[] columns = bodyBox.SplitColumnsCm(2, gutterCm);
            PowerPointLayoutBox leftColumn = columns[0];
            PowerPointLayoutBox rightColumn = columns[1];

            if (File.Exists(imagePath)) {
                slide.SetBackgroundImage(imagePath);
            }

            PowerPointTextBox textBox = slide.AddTitle("Shapes and images", titleBox);
            textBox.FontSize = 30;
            textBox.Color = "1F4E79";

            const double captionGapCm = 0.6;
            const double captionHeightCm = 1.2;
            double canvasHeightCm = leftColumn.HeightCm - captionGapCm - captionHeightCm;
            PowerPointLayoutBox canvasBox = PowerPointLayoutBox.FromCentimeters(
                leftColumn.LeftCm, leftColumn.TopCm, leftColumn.WidthCm, canvasHeightCm);
            PowerPointLayoutBox captionBox = PowerPointLayoutBox.FromCentimeters(
                leftColumn.LeftCm,
                leftColumn.TopCm + canvasHeightCm + captionGapCm,
                leftColumn.WidthCm,
                captionHeightCm);

            PowerPointAutoShape card = slide.AddRectangleCm(
                    canvasBox.LeftCm, canvasBox.TopCm, canvasBox.WidthCm, canvasBox.HeightCm * 0.6,
                    "Hero Card")
                .Fill("E7F7FF")
                .Stroke("007ACC", 2);
            card.FillTransparency = 6;
            card.Rotation = -1.5;

            PowerPointAutoShape accent = slide.AddRectangleCm(
                    canvasBox.LeftCm + 0.7,
                    canvasBox.TopCm + 0.7,
                    canvasBox.WidthCm * 0.55,
                    canvasBox.HeightCm * 0.35,
                    "Accent Panel")
                .Fill("FFF4E5")
                .Stroke("C48A00", 1.5);
            accent.Rotation = 2;
            accent.BringToFront();

            const double badgeSizeCm = 2.1;
            slide.AddEllipseCm(
                    canvasBox.LeftCm + canvasBox.WidthCm - badgeSizeCm - 0.4,
                    canvasBox.TopCm + 0.4,
                    badgeSizeCm,
                    badgeSizeCm,
                    "Badge")
                .Fill("FDEBD0")
                .Stroke("D35400", 1.5);

            PowerPointAutoShape arrow = slide.AddShapeCm(
                    A.ShapeTypeValues.RightArrow,
                    canvasBox.LeftCm + 0.4,
                    canvasBox.TopCm + canvasBox.HeightCm * 0.68,
                    canvasBox.WidthCm * 0.6,
                    canvasBox.HeightCm * 0.25,
                    "Flow Arrow")
                .Fill("D9EAD3")
                .Stroke("6AA84F", 1.5);
            arrow.Rotation = -1;

            PowerPointTextBox label = slide.AddTextBox(
                "Rotation, layering, and named shapes for lookup.",
                captionBox);
            label.FontSize = 14;
            label.Color = "1F4E79";

            if (File.Exists(imagePath)) {
                PowerPointPicture picture = slide.AddPicture(imagePath, rightColumn);
                picture.FitToBox(1200, 800, crop: true);
            } else {
                PowerPointTextBox placeholder = slide.AddTextBox("(image placeholder)", rightColumn);
                placeholder.FillColor = "F3F3F3";
                placeholder.OutlineColor = "CCCCCC";
                placeholder.TextVerticalAlignment = A.TextAnchoringTypeValues.Center;
            }

            PowerPointShape? shape = slide.GetShape("Hero Card");
            if (shape != null) {
                shape.BringToFront();
            }
            Console.WriteLine("Found shape: " + shape?.Name);
            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
