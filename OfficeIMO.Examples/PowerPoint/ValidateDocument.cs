using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates validating a presentation document.
    /// </summary>
    public static class ValidateDocument {
        public static void Example(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Validate document");
            string filePath = Path.Combine(folderPath, "ValidateDocument.pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide();
                Console.WriteLine(presentation.DocumentIsValid);
                Console.WriteLine(presentation.DocumentValidationErrors);
                presentation.Save();
            }
        }
    }
}
