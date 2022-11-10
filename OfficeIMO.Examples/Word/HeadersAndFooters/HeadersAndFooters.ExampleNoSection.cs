using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class HeadersAndFooters {
        internal static void Example_BasicWordWithHeaderAndFooterWithoutSections(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Headers and Footers");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Headers and Footers.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "This is a test for Title";
                document.BuiltinDocumentProperties.Category = "This is a test for Category";

                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;


                var paragraph = document.AddParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                var paragraphInHeaderO = document.Header.Default.AddParagraph();
                paragraphInHeaderO.Text = "Odd Header / Section 0";

                var paragraphInHeaderE = document.Header.Even.AddParagraph();
                paragraphInHeaderE.Text = "Even Header / Section 0";

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                // 2 section, 9 paragraphs + 7 pagebreaks = 15 paragraphs, 7 pagebreaks
                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);

                // primary section (for the whole document)
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
                // additional sections
                document.Save(openWord);
            }
        }

    }
}
