using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal partial class Sections {

        internal static void Example_BasicSections(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with sections");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with some sections.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test 1 - Should be before 1st section").SetColor(Color.LightPink);

                var section1 = document.AddSection();
                section1.AddParagraph("Test 1 - Should be after 1st section").SetFontFamily("Tahoma").SetFontSize(20);

                document.AddParagraph("Test 2 - Should be after 1st section");
                var section2 = document.AddSection();

                document.AddParagraph("Test 3 - Should be after 2nd section");
                document.AddParagraph("Test 4 - Should be after 2nd section").SetBold().AddText(" more text").SetColor(Color.DarkSalmon);

                var section3 = document.AddSection();

                var para = document.AddParagraph("Test 5 -");
                para = para.AddText(" and more text");
                para.Bold = true;

                document.AddPageBreak();

                var paragraph = document.AddParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Blue;

                paragraph.SetBold().SetFontFamily("Tahoma");
                paragraph.AddText(" This is continuation").SetUnderline(UnderlineValues.Double).SetHighlight(HighlightColorValues.DarkGreen).SetFontSize(15).SetColor(Color.Aqua);

                paragraph = document.AddParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Blue;

                paragraph.SetBold().SetFontFamily("Tahoma");
                paragraph.AddText(" This is continuation").SetUnderline(UnderlineValues.Double).SetHighlight(HighlightColorValues.DarkGreen).SetFontSize(15).SetColor(Color.Yellow);


                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);

                // primary section (for the whole document)
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 1: " + document.Sections[1].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 2: " + document.Sections[2].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 3: " + document.Sections[3].Paragraphs.Count);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);

                // primary section (for the whole document)
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 1: " + document.Sections[1].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 2: " + document.Sections[2].Paragraphs.Count);
                Console.WriteLine("+ Paragraphs section 3: " + document.Sections[3].Paragraphs.Count);

                document.Save(openWord);
            }
        }


    }
}
