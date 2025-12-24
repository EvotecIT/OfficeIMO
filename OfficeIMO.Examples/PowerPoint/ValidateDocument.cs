using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates validating a PowerPoint presentation.
    /// </summary>
    public static class ValidateDocument {
        public static void Example(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Validate document");
            string filePath = Path.Combine(folderPath, "ValidateDocument.pptx");
            const double marginCm = 1.5;

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);

            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTitleCm("Validation Example", marginCm, marginCm, content.WidthCm, 1.3);
            slide.AddTextBoxCm("Creates a deck and validates it using Open XML rules.",
                marginCm, 3.2, content.WidthCm, 1.2);

            presentation.Save();
            var errors = presentation.ValidateDocument();

            PowerPointSlide results = presentation.AddSlide();
            results.AddTitleCm("Validation Result", marginCm, marginCm, content.WidthCm, 1.3);
            string summary = errors.Count == 0 ? "No validation errors found." : $"Errors: {errors.Count}";
            results.AddTextBoxCm(summary, marginCm, 3.2, content.WidthCm, 1.0);

            if (errors.Count > 0) {
                PowerPointTextBox details = results.AddTextBoxCm("Top issues:", marginCm, 4.6, content.WidthCm, 2.0);
                foreach (string message in errors.Select(e => e.Description).Take(3)) {
                    details.AddBullet(message);
                }
            }

            presentation.Save();
            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
