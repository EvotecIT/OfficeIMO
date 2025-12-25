using System;
using System.IO;
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
            const double marginCm = 1.5;
            const double gutterCm = 1.0;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(marginCm);
            double bodyTopCm = content.TopCm + 1.9;
            double bodyHeightCm = content.HeightCm - 1.9;
            PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumnsCm(2, marginCm, gutterCm);
            PowerPointLayoutBox leftColumn = new(columns[0].Left, PowerPointUnits.FromCentimeters(bodyTopCm), columns[0].Width,
                PowerPointUnits.FromCentimeters(bodyHeightCm));
            PowerPointLayoutBox rightColumn = new(columns[1].Left, PowerPointUnits.FromCentimeters(bodyTopCm), columns[1].Width,
                PowerPointUnits.FromCentimeters(bodyHeightCm));

            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox title = slide.AddTitleCm("Text Formatting",
                content.LeftCm, content.TopCm, content.WidthCm, 1.4);
            if (title.Paragraphs.Count > 0) {
                PowerPointTextStyle.Title.WithColor("1F4E79").Apply(title.Paragraphs[0]);
            }

            PowerPointTextBox text = slide.AddTextBoxCm(string.Empty,
                leftColumn.LeftCm, leftColumn.TopCm, leftColumn.WidthCm, leftColumn.HeightCm);
            text.SetTextMarginsCm(0.3, 0.2, 0.3, 0.2);
            text.TextAutoFit = PowerPointTextAutoFit.Normal;

            PowerPointParagraph heading = text.AddParagraph("Formatting demo", p => {
                p.Alignment = A.TextAlignmentTypeValues.Left;
                p.SpaceAfterPoints = 6;
            });
            PowerPointTextStyle.Subtitle.WithColor("1F4E79").Apply(heading);

            PowerPointParagraph line = text.AddParagraph();
            line.AddText("Mix ");
            line.AddFormattedText("bold", bold: true).SetColor("C00000");
            line.AddText(", ");
            line.AddFormattedText("italic", italic: true).SetColor("0070C0");
            line.AddText(", and ");
            line.AddFormattedText("underline", underline: A.TextUnderlineValues.Single).SetColor("38761D");
            line.AddText(" runs in one paragraph.");

            text.AddBullet("Bulleted item one");
            text.AddBullet("Bulleted item two");
            text.AddNumberedItem("First step");
            text.AddNumberedItem("Second step");
            text.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);

            PowerPointTextBox callout = slide.AddTextBoxCm(
                "Auto-fit + margins keep text readable.",
                rightColumn.LeftCm, rightColumn.TopCm, rightColumn.WidthCm, rightColumn.HeightCm);
            callout.FillColor = "E7F7FF";
            callout.OutlineColor = "5B9BD5";
            callout.TextVerticalAlignment = A.TextAnchoringTypeValues.Center;
            callout.SetTextMarginsCm(0.3, 0.3, 0.3, 0.3);
            callout.ApplyTextStyle(PowerPointTextStyle.Body.WithColor("1F4E79"));

            presentation.Save();
            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
