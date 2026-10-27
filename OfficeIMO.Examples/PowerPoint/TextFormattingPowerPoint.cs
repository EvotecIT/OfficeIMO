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
            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
            const double marginCm = 1.5;
            const double gutterCm = 1.0;
            PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Centimeters(marginCm));
            double bodyTopCm = content.Top.ToCentimeters() + 1.9;
            double bodyHeightCm = content.Height.ToCentimeters() - 1.9;
            PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumns(2,OfficeIMO.OfficeLength.Centimeters(marginCm),OfficeIMO.OfficeLength.Centimeters(gutterCm));
            PowerPointLayoutBox leftColumn = new(OfficeIMO.OfficeLength.Emus(columns[0].Left.ToEmus()),OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromCentimeters(bodyTopCm)),OfficeIMO.OfficeLength.Emus(columns[0].Width.ToEmus()),                OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromCentimeters(bodyHeightCm)));
            PowerPointLayoutBox rightColumn = new(OfficeIMO.OfficeLength.Emus(columns[1].Left.ToEmus()),OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromCentimeters(bodyTopCm)),OfficeIMO.OfficeLength.Emus(columns[1].Width.ToEmus()),                OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromCentimeters(bodyHeightCm)));

            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox title = slide.AddTitleCm("Text Formatting",
                content.Left.ToCentimeters(), content.Top.ToCentimeters(), content.Width.ToCentimeters(), 1.4);
            if (title.Paragraphs.Count > 0) {
                PowerPointTextStyle.Title.WithColor("1F4E79").Apply(title.Paragraphs[0]);
            }

            PowerPointTextBox text = slide.AddTextBoxCm(string.Empty,
                leftColumn.Left.ToCentimeters(), leftColumn.Top.ToCentimeters(), leftColumn.Width.ToCentimeters(), leftColumn.Height.ToCentimeters());
            text.SetTextMarginsCm(0.3, 0.2, 0.3, 0.2);
            text.TextAutoFit = PowerPointTextAutoFit.Normal;

            PowerPointParagraph heading = text.AddParagraph("Formatting demo", p => {
                p.Alignment = PowerPointTextAlignment.Left;
                p.SpaceAfterPoints = 6;
            });
            PowerPointTextStyle.Subtitle.WithColor("1F4E79").Apply(heading);

            PowerPointParagraph line = text.AddParagraph();
            line.AddText("Mix ");
            line.AddFormattedText("bold", bold: true).SetColor("C00000");
            line.AddText(", ");
            line.AddFormattedText("italic", italic: true).SetColor("0070C0");
            line.AddText(", and ");
            line.AddFormattedText("underline", underline: PowerPointUnderlineStyle.Single).SetColor("38761D");
            line.AddText(" runs in one paragraph.");

            PowerPointParagraph nativeStyles = text.AddParagraph();
            nativeStyles.AddText("Native variants: ");
            PowerPointTextRun decorated = nativeStyles.AddRun("wavy double + double strike");
            decorated.UnderlineStyle = PowerPointUnderlineStyle.WavyDouble;
            decorated.StrikeStyle = PowerPointStrikeStyle.Double;
            decorated.Color = "7030A0";
            nativeStyles.AddText(" H");
            nativeStyles.AddRun("2").SetSubscript();
            nativeStyles.AddText("O and x");
            nativeStyles.AddRun("2").SetSuperscript();
            PowerPointTextRun caps = nativeStyles.AddRun(" small caps");
            caps.Capitalization = PowerPointCapitalization.SmallCaps;
            caps.TransformTextCase(OfficeIMO.Drawing.OfficeTextCase.TitleCase);

            text.AddBullet("Bulleted item one");
            text.AddBullet("Bulleted item two");
            text.AddNumberedItem("First step");
            text.AddNumberedItem("Second step");
            text.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);

            PowerPointTextBox callout = slide.AddTextBoxCm(
                "Auto-fit + margins keep text readable.",
                rightColumn.Left.ToCentimeters(), rightColumn.Top.ToCentimeters(), rightColumn.Width.ToCentimeters(), rightColumn.Height.ToCentimeters());
            callout.FillColor = "E7F7FF";
            callout.OutlineColor = "5B9BD5";
            callout.TextVerticalAlignment = PowerPointTextVerticalAlignment.Center;
            callout.SetTextMarginsCm(0.3, 0.3, 0.3, 0.3);
            callout.ApplyTextStyle(PowerPointTextStyle.Body.WithColor("1F4E79"));

            presentation.Save();
            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }
    }
}
