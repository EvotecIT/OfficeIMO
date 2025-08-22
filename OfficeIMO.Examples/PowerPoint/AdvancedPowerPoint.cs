using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates advanced slide features such as backgrounds, transitions and charts.
    /// </summary>
    public static class AdvancedPowerPoint {
        public static void Example_AdvancedPowerPoint(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Advanced features");
            string filePath = Path.Combine(folderPath, "Advanced PowerPoint.pptx");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox text = slide.AddTextBox("Sample text");
            slide.AddChart();
            slide.BackgroundColor = "FFFF00";
            text.FillColor = "FF0000";
            slide.Transition = SlideTransition.Wipe;
            //slide.Notes.Text = "Demo notes";
            presentation.Save();
        }
    }
}
