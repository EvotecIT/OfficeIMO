using System;
using System.IO;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates theme manipulation and layout selection.
    /// </summary>
    public static class ThemeAndLayoutPowerPoint {
        public static void Example_PowerPointThemeAndLayout(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Theme and Layout presentation");
            string filePath = Path.Combine(folderPath, "ThemeAndLayout.pptx");
            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            const double marginCm = 1.5;
            const double titleHeightCm = 1.6;
            const double bodyGapCm = 0.8;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            PowerPointLayoutBox titleBox = PowerPointLayoutBox.FromCentimeters(
                content.LeftCm, content.TopCm, content.WidthCm, titleHeightCm);
            PowerPointLayoutBox bodyBox = PowerPointLayoutBox.FromCentimeters(
                content.LeftCm,
                content.TopCm + titleHeightCm + bodyGapCm,
                content.WidthCm,
                content.HeightCm - titleHeightCm - bodyGapCm);
            string themeName = string.IsNullOrWhiteSpace(presentation.ThemeName) ? "Office Theme" : presentation.ThemeName;

            PowerPointSlide first = presentation.AddSlide();
            PowerPointTextBox title = first.AddTitle("Theme and Layout", titleBox);
            title.FontSize = 30;
            title.Color = "1F4E79";
            PowerPointLayoutBox summaryBox = PowerPointLayoutBox.FromCentimeters(
                bodyBox.LeftCm, bodyBox.TopCm, bodyBox.WidthCm, 1.1);
            PowerPointTextBox details = first.AddTextBox(
                "Uses the built-in Office theme, layout placeholders, and content boxes.",
                summaryBox);
            details.FontSize = 18;
            details.Color = "1F4E79";
            PowerPointLayoutBox themeBox = PowerPointLayoutBox.FromCentimeters(
                bodyBox.LeftCm, bodyBox.TopCm + 1.6, bodyBox.WidthCm, 0.9);
            first.AddTextBox($"Theme: {themeName}", themeBox);

            PowerPointLayoutBox swatchBox = PowerPointLayoutBox.FromCentimeters(
                bodyBox.LeftCm, bodyBox.TopCm + 3.0, bodyBox.WidthCm, 1.2);
            string[] swatches = { "1F4E79", "5B9BD5", "ED7D31", "A5A5A5", "FFC000" };
            PowerPointLayoutBox[] swatchColumns = swatchBox.SplitColumnsCm(swatches.Length, 0.4);
            for (int i = 0; i < swatches.Length; i++) {
                first.AddRectangleCm(
                        swatchColumns[i].LeftCm,
                        swatchColumns[i].TopCm,
                        swatchColumns[i].WidthCm,
                        swatchColumns[i].HeightCm,
                        $"Swatch {i + 1}")
                    .Fill(swatches[i])
                    .Stroke("FFFFFF", 1);
            }

            PowerPointSlide second = presentation.AddSlide(masterIndex: 0, layoutIndex: 1);
            PowerPointTextBox? titlePlaceholder = second.GetPlaceholder(PlaceholderValues.Title)
                                               ?? second.GetPlaceholder(PlaceholderValues.CenteredTitle);
            if (titlePlaceholder != null) {
                titlePlaceholder.Text = "Layout Placeholders";
                titlePlaceholder.ApplyTextStyle(PowerPointTextStyle.Title.WithColor("1F4E79"));
            } else {
                second.AddTitleCm("Layout Placeholders", marginCm, marginCm, content.WidthCm, titleHeightCm);
            }

            PowerPointTextBox? bodyPlaceholder = second.GetPlaceholder(PlaceholderValues.Body);
            if (bodyPlaceholder != null) {
                bodyPlaceholder.Clear();
                bodyPlaceholder.AddBullets(new[] {
                    "Title and body placeholders from the layout",
                    "Text flows in the layout-defined box",
                    "Margins and spacing stay consistent"
                });
                bodyPlaceholder.SetTextMarginsCm(0.2, 0.1, 0.2, 0.1);
                bodyPlaceholder.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);
            } else {
                second.AddTextBoxCm("Body placeholder not found in layout.", marginCm, 3.4, content.WidthCm, 1.2);
            }

            second.Notes.Text = "Slide created using master/layout indexes and placeholders.";

            // Slide 3: custom layout with content boxes
            PowerPointSlide third = presentation.AddSlide();
            third.AddTitle("Content Box Layout", titleBox);
            PowerPointLayoutBox[] rows = bodyBox.SplitRowsCm(2, 0.6);
            PowerPointLayoutBox[] topColumns = rows[0].SplitColumnsCm(2, 0.6);

            PowerPointTextBox leftCard = third.AddTextBox("Left column\n(overview)", topColumns[0]);
            leftCard.FillColor = "E7F7FF";
            leftCard.OutlineColor = "5B9BD5";
            leftCard.TextVerticalAlignment = A.TextAnchoringTypeValues.Center;
            leftCard.ApplyTextStyle(PowerPointTextStyle.Body.WithColor("1F4E79"));

            PowerPointTextBox rightCard = third.AddTextBox("Right column\n(details)", topColumns[1]);
            rightCard.FillColor = "FFF4E5";
            rightCard.OutlineColor = "C48A00";
            rightCard.TextVerticalAlignment = A.TextAnchoringTypeValues.Center;
            rightCard.ApplyTextStyle(PowerPointTextStyle.Body.WithColor("7F6000"));

            PowerPointTextBox footer = third.AddTextBox("Content boxes keep spacing consistent across slides.", rows[1]);
            footer.FillColor = "F3F3F3";
            footer.OutlineColor = "CCCCCC";
            footer.TextVerticalAlignment = A.TextAnchoringTypeValues.Center;

            presentation.Save();
            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
