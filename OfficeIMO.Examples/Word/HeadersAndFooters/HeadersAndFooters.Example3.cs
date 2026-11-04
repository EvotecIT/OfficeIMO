using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class HeadersAndFooters {

        internal static void Example_BasicWordWithHeaderAndFooter(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Headers and Footers including Sections");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Headers and Footers.docx");


            using (WordDocument document = WordDocument.Create(filePath)) {

                document.Sections[0].PageSetup.Orientation = OfficePageOrientation.Landscape;

                var setup = document.PageSetup;
                var defaultHeader = setup.Header;

                var paragraphInHeader = defaultHeader.AddParagraph();
                paragraphInHeader.Text = "Default Header / Section 0";

                document.AddPageBreak();

                var paragraph = document.AddParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = WordParagraphAlignment.Center;
                paragraph.Color = OfficeIMO.OfficeColor.Red;

                var section2 = document.AddSection();

                var section2Setup = section2.PageSetup;
                var section2DefaultHeader = section2Setup.Header;

                var paragraghInHeaderSection1 = section2DefaultHeader.AddParagraph();
                paragraghInHeaderSection1.Text = "Section 1 header";

                paragraph = document.AddParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = WordParagraphAlignment.Center;
                paragraph.Color = OfficeIMO.OfficeColor.Red;

                var section3 = document.AddSection();

                var section3Setup = section3.PageSetup;
                var section3DefaultHeader = section3Setup.Header;

                var paragraghInHeaderSection3 = section3DefaultHeader.AddParagraph();
                paragraghInHeaderSection3.Text = "Section 2 header";

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
                Console.WriteLine("+ Paragraphs section 1: " + document.Sections[1].Paragraphs.Count);
                //Console.WriteLine("+ Paragraphs section 2: " + document.Sections[0].Paragraphs.Count);
                //Console.WriteLine("+ Paragraphs section 3: " + document.Sections[0].Paragraphs.Count);
                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }

    }
}
