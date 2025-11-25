using System;
using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates text formatting within a textbox.
    /// </summary>
    public static class TextFormattingPowerPoint {
        public static void Example_TextFormattingPowerPoint(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Text formatting");
            string filePath = Path.Combine(folderPath, "Text Formatting PowerPoint.pptx");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox text = slide.AddTextBox("Hello World");
            text.Bold = true;
            text.Italic = true;
            text.FontSize = 24;
            text.FontName = "Arial";
            text.Color = "FF0000";
            text.AddBullet("First bullet");
            text.AddBullet("Second bullet");
            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
