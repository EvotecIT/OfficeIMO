using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
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

            using PresentationDocument document = PresentationDocument.Open(filePath, false);
            var part = document.PresentationPart!;
            Console.WriteLine($"Notes master part present: {part.NotesMasterPart != null}");
            Console.WriteLine($"Presentation properties part present: {part.PresentationPropertiesPart != null}");
            Console.WriteLine($"View properties part present: {part.ViewPropertiesPart != null}");
            Console.WriteLine($"Table styles part present: {part.TableStylesPart != null}");
        }
    }
}
