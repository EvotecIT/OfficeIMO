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
                document.Properties.Title = "This is a test for Title";
                document.Properties.Category = "This is a test for Category";

                document.PageSetup.DifferentOddAndEvenPages = true;



                var paragraph = document.AddParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = WordParagraphAlignment.Center;
                paragraph.Color = OfficeIMO.OfficeColor.Red;

                var setup = document.PageSetup;
                var defaultHeader = setup.Header;
                var evenHeader = setup.EvenPageHeader;

                var paragraphInHeaderO = defaultHeader.AddParagraph();
                paragraphInHeaderO.Text = "Odd Header / Section 0";

                var paragraphInHeaderE = evenHeader.AddParagraph();
                paragraphInHeaderE.Text = "Even Header / Section 0";

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = WordParagraphAlignment.Center;
                paragraph.Color = OfficeIMO.OfficeColor.Red;

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = WordParagraphAlignment.Center;
                paragraph.Color = OfficeIMO.OfficeColor.Red;

                // 2 section, 9 paragraphs + 7 pagebreaks = 15 paragraphs, 7 pagebreaks
                Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
                Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
                Console.WriteLine("+ Sections: " + document.Sections.Count);

                // primary section (for the whole document)
                Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
                // additional sections
                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }

    }
}
