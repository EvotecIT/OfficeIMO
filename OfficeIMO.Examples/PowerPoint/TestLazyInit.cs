using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates creating slides lazily, saving, reopening, and appending content.
    /// </summary>
    public static class TestLazyInit {
        public static void Example_TestLazyInit(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Testing Lazy Initialization");
            string filePath = Path.Combine(folderPath, "Test Lazy Init.pptx");
            const double marginCm = 1.5;
            const double titleHeightCm = 1.3;
            const double bodyGapCm = 0.8;

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                Console.WriteLine($"  Initial slides: {presentation.Slides.Count} (expected: 0)");
                PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
                PowerPointLayoutBox titleBox = PowerPointLayoutBox.FromCentimeters(
                    content.LeftCm, content.TopCm, content.WidthCm, titleHeightCm);
                PowerPointLayoutBox bodyBox = PowerPointLayoutBox.FromCentimeters(
                    content.LeftCm,
                    content.TopCm + titleHeightCm + bodyGapCm,
                    content.WidthCm,
                    content.HeightCm - titleHeightCm - bodyGapCm);

                PowerPointSlide intro = presentation.AddSlide();
                intro.AddTitle("Lifecycle Example", titleBox);
                PowerPointTextBox introBox = intro.AddTextBox(
                    "Slides are created on demand and persisted on Save().",
                    bodyBox);
                introBox.AddBullet("Create slides in memory");
                introBox.AddBullet("Save to persist content");
                introBox.AddBullet("Reopen and append later");
                introBox.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);

                PowerPointSlide step1 = presentation.AddSlide();
                step1.AddTitle("Step 1: Create", titleBox);
                PowerPointTextBox step1Box = step1.AddTextBox(
                    $"Created two slides in memory. Current count: {presentation.Slides.Count}",
                    bodyBox);
                step1Box.TextAutoFit = PowerPointTextAutoFit.Normal;
                step1Box.SetTextMarginsCm(0.3, 0.2, 0.3, 0.2);
                Console.WriteLine($"  After 2 slides: {presentation.Slides.Count} (expected: 2)");

                presentation.Save();
            }

            // Reopen and append more content
            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
                PowerPointLayoutBox titleBox = PowerPointLayoutBox.FromCentimeters(
                    content.LeftCm, content.TopCm, content.WidthCm, titleHeightCm);
                PowerPointLayoutBox bodyBox = PowerPointLayoutBox.FromCentimeters(
                    content.LeftCm,
                    content.TopCm + titleHeightCm + bodyGapCm,
                    content.WidthCm,
                    content.HeightCm - titleHeightCm - bodyGapCm);
                Console.WriteLine($"  Reopened slides: {presentation.Slides.Count} (expected: 2)");

                PowerPointSlide step2 = presentation.AddSlide();
                step2.AddTitle("Step 2: Reopen", titleBox);
                PowerPointTextBox step2Box = step2.AddTextBox(
                    "Added a third slide after reopening the file.",
                    bodyBox);
                step2Box.AddBullet($"Total slides now: {presentation.Slides.Count}");
                step2Box.AddBullet("Changes are persisted on Save()");
                step2Box.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);
                step2.Notes.Text = "Demonstrates lazy initialization and persistence.";

                presentation.Save();
            }

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
