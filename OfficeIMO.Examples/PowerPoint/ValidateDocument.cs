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
                Console.WriteLine(presentation.DocumentValidationErrors);
                presentation.Save();
            }

            if (openPowerPoint) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

