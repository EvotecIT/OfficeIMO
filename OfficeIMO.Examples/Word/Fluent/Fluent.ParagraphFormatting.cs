using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentParagraphFormatting(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with formatted paragraphs using fluent API");
            string filePath = Path.Combine(folderPath, "FluentParagraphFormatting.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraph(p => p.Text("Left aligned paragraph").AlignLeft())
                    .Paragraph(p => p.Text("Centered paragraph").AlignCenter())
                    .Paragraph(p => p.Text("Right aligned paragraph").AlignRight())
                    .Paragraph(p => p.Text("Justified heading with spacing and indentation")
                        .AlignJustified()
                        .SpacingBefore(12)
                        .SpacingAfter(12)
                        .LineSpacing(24)
                        .Indentation(left: 24, firstLine: 24)
                        .Style(WordParagraphStyles.Heading2))
                    .Paragraph(p => p.Text("Bullet list item").AddList(WordListStyle.Bulleted))
                    .Paragraph(p => p.Text("Table below").AddTableAfter(2, 2))
                    .End();
                document.Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
