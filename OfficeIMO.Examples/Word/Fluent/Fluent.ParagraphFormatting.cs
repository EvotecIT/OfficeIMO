using System;
using System.IO;
using OfficeIMO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        public static void Example_FluentParagraphFormatting(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with formatted paragraphs using fluent API");
            string filePath = Path.Combine(folderPath, "FluentParagraphFormatting.docx");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    var paragraph1 = document.AddParagraph();
                    paragraph1 = paragraph1.AddText("Left aligned paragraph");
                    paragraph1.SetAlignment(WordParagraphAlignment.Left);
                    var paragraph2 = document.AddParagraph();
                    paragraph2 = paragraph2.AddText("Centered paragraph");
                    paragraph2.SetAlignment(WordParagraphAlignment.Center);
                    var paragraph3 = document.AddParagraph();
                    paragraph3 = paragraph3.AddText("Right aligned paragraph");
                    paragraph3.SetAlignment(WordParagraphAlignment.Right);
                    var paragraph4 = document.AddParagraph();
                    paragraph4 = paragraph4.AddText("Justified heading with spacing and indentation");
                    paragraph4.SetAlignment(WordParagraphAlignment.Both);
                    paragraph4.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(12D);
                    paragraph4.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(12D);
                    paragraph4.LineSpacing = (int)(24 * 20);
                    paragraph4.IndentationBefore = global::OfficeIMO.OfficeLength.Points(((int)(24 * 20)) / 20D);
                    paragraph4.IndentationFirstLine = global::OfficeIMO.OfficeLength.Points(((int)(24 * 20)) / 20D);
                    paragraph4.SetStyle(WordParagraphStyles.Heading2);
                    var list5 = document.AddList(WordListStyle.Bulleted);
                    int level6 = 0;
                    list5.AddItem("Bullet list item", level6);
                    var paragraph7 = document.AddParagraph();
                    paragraph7 = paragraph7.AddText("Table below");
                    var table8 = document.AddTable(2, 2);
                    document.Save();
                }
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}