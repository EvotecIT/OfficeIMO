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

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                Console.WriteLine($"  Initial slides: {presentation.Slides.Count} (expected: 0)");
                PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);

                PowerPointSlide intro = presentation.AddSlide();
                intro.AddTitleCm("Lifecycle Example", marginCm, marginCm, content.WidthCm, 1.3);
                intro.AddTextBoxCm("Slides are created on demand and persisted on Save().",
                    marginCm, 3.2, content.WidthCm, 1.2);

                PowerPointSlide step1 = presentation.AddSlide();
                step1.AddTitleCm("Step 1: Create", marginCm, marginCm, content.WidthCm, 1.3);
                step1.AddTextBoxCm("Created two slides in memory.", marginCm, 3.2, content.WidthCm, 1.0);
                Console.WriteLine($"  After 2 slides: {presentation.Slides.Count} (expected: 2)");

                presentation.Save();
            }

            // Reopen and append more content
            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
                Console.WriteLine($"  Reopened slides: {presentation.Slides.Count} (expected: 2)");

                PowerPointSlide step2 = presentation.AddSlide();
                step2.AddTitleCm("Step 2: Reopen", marginCm, marginCm, content.WidthCm, 1.3);
                step2.AddTextBoxCm("Added a third slide after reopening.", marginCm, 3.2, content.WidthCm, 1.0);
                step2.AddTextBoxCm($"Total slides: {presentation.Slides.Count}", marginCm, 4.6, content.WidthCm, 1.0);
                step2.Notes.Text = "Demonstrates lazy initialization and persistence.";

                presentation.Save();
            }

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
