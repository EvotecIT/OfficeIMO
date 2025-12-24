using System;
using System.IO;
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

            presentation.ThemeName = "OfficeIMO Theme";
            const double marginCm = 1.5;
            const double gutterCm = 1.0;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);

            PowerPointSlide first = presentation.AddSlide();
            first.AddTitleCm("Theme and Layout", marginCm, marginCm, content.WidthCm, 1.3);
            PowerPointTextBox details = first.AddTextBoxCm("Theme name and layout selection are explicit.",
                marginCm, 3.2, content.WidthCm, 1.2);
            details.FontSize = 18;
            details.Color = "1F4E79";
            first.AddTextBoxCm($"Theme: {presentation.ThemeName}", marginCm, 4.6, content.WidthCm, 1.0);

            PowerPointSlide second = presentation.AddSlide(masterIndex: 0, layoutIndex: 0);
            second.AddTitleCm("Layout Slots", marginCm, marginCm, content.WidthCm, 1.3);
            PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumnsCm(2, marginCm, gutterCm);
            PowerPointTextBox left = second.AddTextBox("Left column", columns[0].Left, columns[0].Top, columns[0].Width,
                PowerPointUnits.FromCentimeters(1.4));
            left.FontSize = 16;
            PowerPointTextBox right = second.AddTextBox("Right column", columns[1].Left, columns[1].Top, columns[1].Width,
                PowerPointUnits.FromCentimeters(1.4));
            right.FontSize = 16;
            second.Notes.Text = "Slide created using master/layout indexes.";

            presentation.Save();
            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
