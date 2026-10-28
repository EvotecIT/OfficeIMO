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
            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
            const double marginCm = 1.5;
            const double titleHeightCm = 1.6;
            const double bodyGapCm = 0.8;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Centimeters(marginCm));
            PowerPointLayoutBox titleBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(titleHeightCm));
            PowerPointLayoutBox bodyBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),                OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + titleHeightCm + bodyGapCm),                OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),                OfficeIMO.OfficeLength.Centimeters(content.Height.ToCentimeters() - titleHeightCm - bodyGapCm));
            string themeName = string.IsNullOrWhiteSpace(presentation.ThemeName) ? "Office Theme" : presentation.ThemeName;

            PowerPointSlide first = presentation.AddSlide();
            PowerPointTextBox title = first.AddTitle("Theme and Layout",titleBox);
            title.FontSize = global::OfficeIMO.OfficeLength.Points(30);
            title.Color = "1F4E79";
            PowerPointLayoutBox summaryBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(bodyBox.Left.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(bodyBox.Top.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(bodyBox.Width.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(1.1));
            PowerPointTextBox details = first.AddTextBox(
                "Uses the built-in Office theme, layout placeholders, and content boxes.",                summaryBox);
            details.FontSize = global::OfficeIMO.OfficeLength.Points(18);
            details.Color = "1F4E79";
            PowerPointLayoutBox themeBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(bodyBox.Left.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(bodyBox.Top.ToCentimeters() + 1.6),OfficeIMO.OfficeLength.Centimeters(bodyBox.Width.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(0.9));
            first.AddTextBox($"Theme: {themeName}",themeBox);

            PowerPointLayoutBox swatchBox = new OfficeIMO.PowerPoint.PowerPointLayoutBox(
                OfficeIMO.OfficeLength.Centimeters(bodyBox.Left.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(bodyBox.Top.ToCentimeters() + 3.0),OfficeIMO.OfficeLength.Centimeters(bodyBox.Width.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(1.2));
            string[] swatches = { "1F4E79", "5B9BD5", "ED7D31", "A5A5A5", "FFC000" };
            PowerPointLayoutBox[] swatchColumns = swatchBox.SplitColumns(swatches.Length,OfficeIMO.OfficeLength.Centimeters(0.4));
            for (int i = 0; i < swatches.Length; i++) {
                first.AddRectangle(
                        swatchColumns[i].Left,
                        swatchColumns[i].Top,
                        swatchColumns[i].Width,
                        swatchColumns[i].Height,
                        $"Swatch {i + 1}")
                    .Fill(swatches[i])
                    .Stroke("FFFFFF", 1);
            }

            PowerPointSlide second = presentation.AddSlide(masterIndex: 0, layoutIndex: 1);
            PowerPointTextBox? titlePlaceholder = second.GetPlaceholder(PowerPointPlaceholderType.Title)
                                               ?? second.GetPlaceholder(PowerPointPlaceholderType.CenteredTitle);
            if (titlePlaceholder != null) {
                titlePlaceholder.Text = "Layout Placeholders";
                titlePlaceholder.ApplyTextStyle(PowerPointTextStyle.Title.WithColor("1F4E79"));
            } else {
                second.AddTitle("Layout Placeholders",global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(titleHeightCm));
            }

            PowerPointTextBox? bodyPlaceholder = second.GetPlaceholder(PowerPointPlaceholderType.Body);
            if (bodyPlaceholder != null) {
                bodyPlaceholder.Clear();
                bodyPlaceholder.AddBullets(new[] {
                    "Title and body placeholders from the layout",
                    "Text flows in the layout-defined box",
                    "Margins and spacing stay consistent"
                });
                bodyPlaceholder.SetTextMargins(global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.1), global::OfficeIMO.OfficeLength.Centimeters(0.2), global::OfficeIMO.OfficeLength.Centimeters(0.1));
                bodyPlaceholder.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfter: global::OfficeIMO.OfficeLength.Points(2));
            } else {
                second.AddTextBox("Body placeholder not found in layout.",global::OfficeIMO.OfficeLength.Centimeters(marginCm),global::OfficeIMO.OfficeLength.Centimeters(3.4),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.2));
            }

            second.Notes.Text = "Slide created using master/layout indexes and placeholders.";

            // Slide 3: custom layout with content boxes
            PowerPointSlide third = presentation.AddSlide();
            third.AddTitle("Content Box Layout",titleBox);
            PowerPointLayoutBox[] rows = bodyBox.SplitRows(2,OfficeIMO.OfficeLength.Centimeters(0.6));
            PowerPointLayoutBox[] topColumns = rows[0].SplitColumns(2,OfficeIMO.OfficeLength.Centimeters(0.6));

            PowerPointTextBox leftCard = third.AddTextBox("Left column\n(overview)",topColumns[0]);
            leftCard.FillColor = "E7F7FF";
            leftCard.OutlineColor = "5B9BD5";
            leftCard.TextVerticalAlignment = PowerPointTextVerticalAlignment.Center;
            leftCard.ApplyTextStyle(PowerPointTextStyle.Body.WithColor("1F4E79"));

            PowerPointTextBox rightCard = third.AddTextBox("Right column\n(details)",topColumns[1]);
            rightCard.FillColor = "FFF4E5";
            rightCard.OutlineColor = "C48A00";
            rightCard.TextVerticalAlignment = PowerPointTextVerticalAlignment.Center;
            rightCard.ApplyTextStyle(PowerPointTextStyle.Body.WithColor("7F6000"));

            PowerPointTextBox footer = third.AddTextBox("Content boxes keep spacing consistent across slides.",rows[1]);
            footer.FillColor = "F3F3F3";
            footer.OutlineColor = "CCCCCC";
            footer.TextVerticalAlignment = PowerPointTextVerticalAlignment.Center;

            presentation.Save();
            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }
    }
}
