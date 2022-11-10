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
                document.AddHeadersAndFooters();

                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;

                var paragraphInHeader = document.Header.Default.AddParagraph();
                paragraphInHeader.Text = "Default Header / Section 0";

                document.AddPageBreak();

                var paragraph = document.AddParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                var section2 = document.AddSection();
                section2.AddHeadersAndFooters();

                var paragraghInHeaderSection1 = section2.Header.Default.AddParagraph();
                paragraghInHeaderSection1.Text = "Weird shit? 1";

                paragraph = document.AddParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                var section3 = document.AddSection();
                section3.AddHeadersAndFooters();

                var paragraghInHeaderSection3 = section3.Header.Default.AddParagraph();
                paragraghInHeaderSection3.Text = "Weird shit? 2";

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
                Console.WriteLine("+ Paragraphs section 1: " + document.Sections[1].Paragraphs.Count);
                //Console.WriteLine("+ Paragraphs section 2: " + document.Sections[0].Paragraphs.Count);
                //Console.WriteLine("+ Paragraphs section 3: " + document.Sections[0].Paragraphs.Count);
                document.Save(openWord);
            }
        }

    }
}
