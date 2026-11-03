using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class HeadersAndFooters {
        internal static void Sections1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Sections - Headers/Footers");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Sections - HeadersAndFooters.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Basic paragraph");
                document.AddTable(1, 1);

                var para = document.AddBookmark("Test").Paragraph;

                // lets add some empty space so we can test if bookmark works properly
                document.AddPageBreak();
                document.AddPageBreak();
                document.AddPageBreak();

                var section0 = document.Sections[0];
                var section0Setup = section0.PageSetup;
                var section0DefaultHeader = section0Setup.Header;

                section0DefaultHeader.AddParagraph().AddText("Section 0").AddBookmark("BookmarkInSection0Header1");

                var tableHeader = section0DefaultHeader.AddTable(3, 4);
                tableHeader.Rows[0].Cells[3].Paragraphs[0].Text = "This is sparta";
                Console.WriteLine(section0DefaultHeader.Tables.Count);

                section0DefaultHeader.AddHorizontalLine();

                section0DefaultHeader.AddHyperLink("Link to website!", new Uri("https://evotec.xyz"));

                section0DefaultHeader.AddHyperLink("Przemysław Klys Email Me", new Uri("mailto:kontakt@evotec.pl?subject=Test Subject"));

                section0DefaultHeader.AddField(WordFieldType.Author, WordFieldFormat.FirstCap);


                var section0DefaultFooter = section0Setup.Footer;

                section0DefaultFooter.AddParagraph().AddText("Section 0").AddBookmark("BookmarkInSection0Header2");

                var tableFooter = section0DefaultFooter.AddTable(2, 3);
                tableFooter.Rows[0].Cells[2].Paragraphs[0].Text = "This is not sparta";

                section0DefaultFooter.AddHorizontalLine();

                section0DefaultFooter.AddHyperLink("Link to website!", new Uri("https://evotec.xyz"));

                section0DefaultFooter.AddHyperLink("Przemysław Klys Email Me", new Uri("mailto:kontakt@evotec.pl?subject=Test Subject"));

                section0DefaultFooter.AddField(WordFieldType.Author, WordFieldFormat.FirstCap);


                var section1 = document.AddSection();
                section1.AddParagraph("Test Middle1 Section - 1");
                var section1Setup = section1.PageSetup;
                var section1DefaultHeader = section1Setup.Header;
                section1DefaultHeader.AddParagraph().AddText("Section 1 - Header");
                var section1DefaultFooter = section1Setup.Footer;
                section1DefaultFooter.AddParagraph().AddText("Section 1 - Footer");

                var section2 = document.AddSection();
                section2.AddParagraph("Test Middle2 Section - 1");
                var section2Setup = section2.PageSetup;
                var section2DefaultHeader = section2Setup.Header;
                section2DefaultHeader.AddParagraph().AddText("Section 2 - Header");
                var section2DefaultFooter = section2Setup.Footer;
                section2DefaultFooter.AddParagraph().AddText("Section 2 - Footer");

                var section3 = document.AddSection();
                section3.AddParagraph("Test Last Section - 1");
                section3.PageSetup.DifferentOddAndEvenPages = true;
                section3.PageSetup.DifferentFirstPage = true;
                var section3Setup = section3.PageSetup;
                var section3DefaultHeader = section3Setup.Header;
                section3DefaultHeader.AddParagraph().AddText("Section 3 - Header Odd/Default");
                var section3DefaultFooter = section3Setup.Footer;
                section3DefaultFooter.AddParagraph().AddText("Section 3 - Footer Odd/Default");
                var section3EvenHeader = section3Setup.EvenPageHeader;
                section3EvenHeader.AddParagraph().AddText("Section 3 - Header Even");
                var section3EvenFooter = section3Setup.EvenPageFooter;
                section3EvenFooter.AddParagraph().AddText("Section 3 - Footer Even");

                document.AddPageBreak();
                section3.AddParagraph("Test Last Section - 2");
                document.AddPageBreak();
                section3.AddParagraph("Test Last Section - 3");

                document.Save();
                if (openWord) document.OpenInApplication();

                Console.WriteLine("IsValid: " + document.DocumentIsValid);

                Console.WriteLine("Section 0 DifferentOddAndEventPages: " + document.Sections[0].PageSetup.DifferentOddAndEvenPages);
                Console.WriteLine("Section 0 DifferentFirstPage: " + document.Sections[0].PageSetup.DifferentFirstPage);

                Console.WriteLine("Section 1 DifferentOddAndEventPages: " + document.Sections[1].PageSetup.DifferentOddAndEvenPages);
                Console.WriteLine("Section 1 DifferentFirstPage: " + document.Sections[1].PageSetup.DifferentFirstPage);

                Console.WriteLine("Section 2 DifferentOddAndEventPages: " + document.Sections[2].PageSetup.DifferentOddAndEvenPages);
                Console.WriteLine("Section 2 DifferentFirstPage: " + document.Sections[2].PageSetup.DifferentFirstPage);

                Console.WriteLine("Section 3 DifferentOddAndEventPages: " + document.Sections[3].PageSetup.DifferentOddAndEvenPages);
                Console.WriteLine("Section 3 DifferentFirstPage: " + document.Sections[3].PageSetup.DifferentFirstPage);

            }
        }
    }
}
