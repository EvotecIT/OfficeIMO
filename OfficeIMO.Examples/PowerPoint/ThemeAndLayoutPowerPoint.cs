using System;
using System.IO;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;

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
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            string themeName = string.IsNullOrWhiteSpace(presentation.ThemeName) ? "Office Theme" : presentation.ThemeName;

            PowerPointSlide first = presentation.AddSlide();
            PowerPointTextBox title = first.AddTitleCm("Theme and Layout", marginCm, marginCm, content.WidthCm, titleHeightCm);
            title.FontSize = 30;
            title.Color = "1F4E79";
            PowerPointTextBox details = first.AddTextBoxCm("Uses the built-in Office theme, layout placeholders, and content boxes.",
                marginCm, 3.1, content.WidthCm, 1.1);
            details.FontSize = 18;
            details.Color = "1F4E79";
            first.AddTextBoxCm($"Theme: {themeName}", marginCm, 4.5, content.WidthCm, 0.9);

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

            presentation.Save();
            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
