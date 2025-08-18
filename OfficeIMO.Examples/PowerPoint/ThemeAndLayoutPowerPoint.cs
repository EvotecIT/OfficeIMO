using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates theme manipulation and layout selection.
    /// </summary>
    public static class ThemeAndLayoutPowerPoint {
        public static void Example_PowerPointThemeAndLayout(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Theme and Layout presentation");
            string filePath = Path.Combine(folderPath, "ThemeAndLayout.pptx");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            presentation.ThemeName = "Custom Theme";
            PowerPointSlide first = presentation.AddSlide();
            first.AddTextBox("Default layout");
            PowerPointSlide second = presentation.AddSlide(layoutIndex: 1);
            second.AddTextBox("Second layout");
            presentation.Save();
        }
    }
}
