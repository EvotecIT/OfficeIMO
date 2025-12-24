using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;

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
            PowerPointTextBox text = slide.AddTextBox(string.Empty, 914400, 914400, 9144000, 3657600);

            PowerPointParagraph heading = text.AddParagraph("Text formatting demo", p => {
                p.Alignment = A.TextAlignmentTypeValues.Center;
                p.SpaceAfterPoints = 6;
            });
            var headingRun = heading.Runs.First();
            headingRun.Bold = true;
            headingRun.FontSize = 32;
            headingRun.Color = "1F4E79";

            PowerPointParagraph line = text.AddParagraph();
            line.AddRun("This line shows ");
            line.AddRun("bold", r => {
                r.Bold = true;
                r.Color = "C00000";
            });
            line.AddRun(" and ");
            line.AddRun("italic", r => {
                r.Italic = true;
                r.Color = "0070C0";
            });
            line.AddRun(" runs.");

            text.AddBullet("Bulleted item one");
            text.AddBullet("Bulleted item two");
            text.AddNumberedItem("First step");
            text.AddNumberedItem("Second step");
            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}

