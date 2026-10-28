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
            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
            const double marginCm = 1.5;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Centimeters(marginCm));

            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox title = slide.AddTitle("Defaults & Theme",                global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.4));
            if (title.Paragraphs.Count > 0) {
                PowerPointTextStyle.Title.WithColor("1F4E79").Apply(title.Paragraphs[0]);
            }

            string summary = $"Theme: {presentation.ThemeName}\n" +
                             $"Slide size: {presentation.SlideSize.Width.ToCentimeters():0.0} cm × {presentation.SlideSize.Height.ToCentimeters():0.0} cm";
            PowerPointTextBox body = slide.AddTextBox(summary,                global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + 1.9),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(1.6));
            body.ApplyTextStyle(PowerPointTextStyle.Body);
            body.SetTextMarginsCm(0.2, 0.2, 0.2, 0.2);

            PowerPointTextBox notes = slide.AddTextBox(string.Empty,                global::OfficeIMO.OfficeLength.Centimeters(content.Left.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Top.ToCentimeters() + 3.6),global::OfficeIMO.OfficeLength.Centimeters(content.Width.ToCentimeters()),global::OfficeIMO.OfficeLength.Centimeters(content.Height.ToCentimeters() - 3.6));
            notes.SetTextMarginsCm(0.25, 0.2, 0.25, 0.2);
            notes.AddBullets(new[] {
                "Default theme and table styles are embedded",
                "Slides start with the standard Office layout",
                "You can customize theme, transitions, and layout later"
            });
            notes.ApplyAutoSpacing(lineSpacingMultiplier: 1.1, spaceAfterPoints: 2);

            presentation.Save();
            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }
    }
}
