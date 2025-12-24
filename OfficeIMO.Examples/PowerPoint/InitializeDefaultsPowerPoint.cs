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
            presentation.AddSlide();
            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}

