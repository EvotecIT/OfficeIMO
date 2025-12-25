using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates basic <see cref="PowerPointPresentation"/> usage.
    /// </summary>
    public static class BasicPowerPointDocument {
        public static void Example_BasicPowerPoint(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Creating basic presentation");
            string filePath = Path.Combine(folderPath, "Basic PowerPoint.pptx");
            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            const double marginCm = 1.5;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);

            Console.WriteLine($"Theme: {presentation.ThemeName}");
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox title = slide.AddTitleCm("OfficeIMO PowerPoint Basics",
                content.LeftCm, content.TopCm, content.WidthCm, 1.4);
            if (title.Paragraphs.Count > 0) {
                PowerPointTextStyle.Title.WithColor("1F4E79").Apply(title.Paragraphs[0]);
            }

            PowerPointTextBox intro = slide.AddTextBoxCm(
                "Create clean .pptx files with simple, readable APIs.",
                content.LeftCm, content.TopCm + 1.8, content.WidthCm, 1.1);
            intro.ApplyTextStyle(PowerPointTextStyle.Body);
            intro.SetTextMarginsCm(0.2, 0.1, 0.2, 0.1);
            intro.TextAutoFit = PowerPointTextAutoFit.Normal;

            PowerPointTextBox agenda = slide.AddTextBoxCm(string.Empty,
                content.LeftCm, content.TopCm + 3.1, content.WidthCm, content.HeightCm - 3.1);
            agenda.SetTextMarginsCm(0.3, 0.2, 0.3, 0.2);
            agenda.AddBullets(new[] {
                "Add slides, titles, and text boxes",
                "Work in centimeters, inches, or points",
                "Build tables and charts from data",
                "Use themes, transitions, and speaker notes"
            });
            agenda.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);

            presentation.Save();
            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
