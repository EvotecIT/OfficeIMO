using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates validating a <see cref="PowerPointPresentation"/>.
    /// </summary>
    public static class ValidatePowerPointDocument {
        public static void Example_ValidatePowerPoint(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Validating presentation");
            string filePath = Path.Combine(folderPath, "Validate PowerPoint.pptx");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            presentation.AddSlide();
            presentation.Save();

            var errors = presentation.ValidatePresentation();
            Console.WriteLine($"Validation errors: {errors.Count}");
        }
    }
}
