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
            PowerPointTextBox text = slide.AddTextBox("Title Slide");
            text.AddBullet("Subtitle");
            slide.BackgroundColor = "FFFFFF";
            slide.Transition = SlideTransition.Wipe;
            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
