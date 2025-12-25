using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates creation of a presentation with default parts and a single slide.
    /// </summary>
    public static class InitializeDefaultsPowerPoint {
        public static void Example_PowerPointInitializeDefaults(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Initialize defaults presentation");
            string filePath = Path.Combine(folderPath, "InitializeDefaults.pptx");
            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            const double marginCm = 1.5;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);

            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox title = slide.AddTitleCm("Defaults & Theme",
                content.LeftCm, content.TopCm, content.WidthCm, 1.4);
            if (title.Paragraphs.Count > 0) {
                PowerPointTextStyle.Title.WithColor("1F4E79").Apply(title.Paragraphs[0]);
            }

            string summary = $"Theme: {presentation.ThemeName}\n" +
                             $"Slide size: {presentation.SlideSize.WidthCm:0.0} cm × {presentation.SlideSize.HeightCm:0.0} cm";
            PowerPointTextBox body = slide.AddTextBoxCm(summary,
                content.LeftCm, content.TopCm + 1.9, content.WidthCm, 1.6);
            body.ApplyTextStyle(PowerPointTextStyle.Body);
            body.SetTextMarginsCm(0.2, 0.2, 0.2, 0.2);

            PowerPointTextBox notes = slide.AddTextBoxCm(string.Empty,
                content.LeftCm, content.TopCm + 3.6, content.WidthCm, content.HeightCm - 3.6);
            notes.SetTextMarginsCm(0.25, 0.2, 0.25, 0.2);
            notes.AddBullets(new[] {
                "Default theme and table styles are embedded",
                "Slides start with the standard Office layout",
                "You can customize theme, transitions, and layout later"
            });
            notes.ApplyAutoSpacing(lineSpacingMultiplier: 1.1, spaceAfterPoints: 2);

            presentation.Save();
            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
