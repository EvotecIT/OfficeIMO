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
            const double titleHeightCm = 1.3;
            const double titleGapCm = 0.4;
            second.AddTitleCm("Layout Slots", marginCm, marginCm, content.WidthCm, titleHeightCm);
            double bodyTopCm = content.TopCm + titleHeightCm + titleGapCm;
            double bodyHeightCm = content.HeightCm - titleHeightCm - titleGapCm;
            PowerPointLayoutBox body = PowerPointLayoutBox.FromCentimeters(content.LeftCm, bodyTopCm, content.WidthCm,
                bodyHeightCm);
            PowerPointLayoutBox[] columns = body.SplitColumnsCm(2, gutterCm);
            PowerPointTextBox left = second.AddTextBoxCm("Left column\n- Layout aware\n- Consistent margins",
                columns[0].LeftCm, columns[0].TopCm, columns[0].WidthCm, 2.2);
            left.FontSize = 16;
            PowerPointTextBox right = second.AddTextBoxCm("Right column\n- Gutter aware\n- Balanced width",
                columns[1].LeftCm, columns[1].TopCm, columns[1].WidthCm, 2.2);
            right.FontSize = 16;
            second.Notes.Text = "Slide created using master/layout indexes.";

            presentation.Save();
            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
