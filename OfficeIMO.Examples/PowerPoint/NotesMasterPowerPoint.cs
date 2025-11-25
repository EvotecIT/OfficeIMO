using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates that notes masters are referenced correctly when creating a presentation.
    /// </summary>
    public static class NotesMasterPowerPoint {
        public static void Example_PowerPointNotesMaster(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Notes master");
            string filePath = Path.Combine(folderPath, "NotesMaster.pptx");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            PowerPointSlide slide = presentation.AddSlide();
            slide.Notes.Text = "Example notes";
            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
