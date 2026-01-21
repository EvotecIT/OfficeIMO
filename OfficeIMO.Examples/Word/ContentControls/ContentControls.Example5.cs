using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class ContentControls {
        internal static void Example_FormattedContentControls(string folderPath, bool openWord) {
            Console.WriteLine("[*] Content controls with formatting");
            string filePath = Path.Combine(folderPath, "DocumentContentControlsFormatting.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var para1 = document.AddParagraph("Formatted (properties): ");
                var sdt1 = para1.AddStructuredDocumentTag(text: "Property styled", alias: "AliasProps", tag: "TagProps");
                sdt1.Bold = true;
                sdt1.Italic = true;
                sdt1.Underline = UnderlineValues.Single;
                sdt1.FontFamily = "Calibri";
                sdt1.FontSize = 12;
                sdt1.ColorHex = "2F5597";
                sdt1.Highlight = HighlightColorValues.Yellow;

                var para2 = document.AddParagraph("Formatted (fluent): ");
                para2.AddStructuredDocumentTag(text: "Fluent styled", alias: "AliasFluent", tag: "TagFluent")
                    .SetBold()
                    .SetItalic()
                    .SetUnderline(UnderlineValues.Single)
                    .SetFontFamily("Calibri")
                    .SetFontSize(14)
                    .SetColorHex("C00000")
                    .SetHighlight(HighlightColorValues.LightGray);

                document.Save(openWord);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var props = Guard.NotNull(document.GetStructuredDocumentTagByTag("TagProps"), "Structured document tag 'TagProps' was not found.");
                Console.WriteLine($"Props control: {props.Text} bold={props.Bold} font={props.FontFamily} size={props.FontSize}");

                var fluent = Guard.NotNull(document.GetStructuredDocumentTagByTag("TagFluent"), "Structured document tag 'TagFluent' was not found.");
                fluent.SetUnderline(UnderlineValues.Double)
                    .SetColorHex("0070C0");

                document.Save(openWord);
            }
        }
    }
}
