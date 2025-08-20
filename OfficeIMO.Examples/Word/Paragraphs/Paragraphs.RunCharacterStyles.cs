using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word;

internal static partial class Paragraphs {

        internal static void Example_RunCharacterStylesSimple(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with character styled paragraphs");
            string filePath = System.IO.Path.Combine(folderPath, "ParagraphRunStylesSimple.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Styled paragraph");
                paragraph.SetCharacterStyle(WordCharacterStyles.Heading1Char);
                paragraph.AddText(" with hyperlink style").SetCharacterStyleId("Hyperlink");
                document.Save(openWord);
            }
        }

        internal static void Example_RunCharacterStylesAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with all character styles");
            string filePath = System.IO.Path.Combine(folderPath, "ParagraphRunStylesAdvanced.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                foreach (WordCharacterStyles style in Enum.GetValues(typeof(WordCharacterStyles))) {
                    var para = document.AddParagraph(style.ToString());
                    para.SetCharacterStyle(style);
                }
                var advanced = document.AddParagraph("Link example");
                var run = advanced.AddText(" visit website");
                run.SetCharacterStyleId("Hyperlink");
                document.Save(openWord);
            }
        }
}
