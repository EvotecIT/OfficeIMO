using System;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates validating a PowerPoint presentation.
    /// </summary>
    public static class ValidateDocument {
        public static void Example(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Validate document");
            string filePath = System.IO.Path.Combine(folderPath, "ValidateDocument.pptx");

            using (var presentation = PowerPointPresentation.Create(filePath)) {
                Console.WriteLine(presentation.DocumentIsValid);
                foreach (var error in presentation.DocumentValidationErrors) {
                    Console.WriteLine($"Validation Error: {error.Description}");
                }
                presentation.Save();
            }

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}

