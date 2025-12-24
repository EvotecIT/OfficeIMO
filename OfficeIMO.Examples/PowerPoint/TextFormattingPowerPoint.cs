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
            PowerPointTextStyle.Title.WithColor("1F4E79").Apply(heading);

            PowerPointParagraph line = text.AddParagraph();
            line.AddText("This line shows ");
            line.AddFormattedText("bold", bold: true).SetColor("C00000");
            line.AddText(" and ");
            line.AddFormattedText("italic", italic: true).SetColor("0070C0");
            line.AddText(" runs.");

            text.AddBullet("Bulleted item one");
            text.AddBullet("Bulleted item two");
            text.AddNumberedItem("First step");
            text.AddNumberedItem("Second step");
            text.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);
            presentation.Save();

            Helpers.Open(filePath, openPowerPoint);
        }
    }
}

