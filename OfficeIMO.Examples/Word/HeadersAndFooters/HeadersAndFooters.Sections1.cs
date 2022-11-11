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

                var para = document.AddBookmark("Test");

                // lets add some empty space so we can test if bookmark works properly
                document.AddPageBreak();
                document.AddPageBreak();
                document.AddPageBreak();

                document.AddHeadersAndFooters();
                document.Sections[0].Header.Default.AddParagraph().AddText("Section 0").AddBookmark("BookmarkInSection0Header1");

                var tableHeader = document.Sections[0].Header.Default.AddTable(3, 4);
                tableHeader.Rows[0].Cells[3].Paragraphs[0].Text = "This is sparta";
                Console.WriteLine(document.Sections[0].Header.Default.Tables.Count);

                document.Sections[0].Header.Default.AddHorizontalLine();

                document.Sections[0].Header.Default.AddHyperLink("Link to website!", new Uri("https://evotec.xyz"));

                document.Sections[0].Header.Default.AddHyperLink("Przemysław Klys Email Me", new Uri("mailto:kontakt@evotec.pl?subject=Test Subject"));

                document.Sections[0].Header.Default.AddField(WordFieldType.Author, WordFieldFormat.FirstCap);


                document.Sections[0].Footer.Default.AddParagraph().AddText("Section 0").AddBookmark("BookmarkInSection0Header2");

                var tableFooter = document.Sections[0].Footer.Default.AddTable(2, 3);
                tableFooter.Rows[0].Cells[2].Paragraphs[0].Text = "This is not sparta";

                document.Sections[0].Footer.Default.AddHorizontalLine();

                document.Sections[0].Footer.Default.AddHyperLink("Link to website!", new Uri("https://evotec.xyz"));

                document.Sections[0].Footer.Default.AddHyperLink("Przemysław Klys Email Me", new Uri("mailto:kontakt@evotec.pl?subject=Test Subject"));

                document.Sections[0].Footer.Default.AddField(WordFieldType.Author, WordFieldFormat.FirstCap);


                var section1 = document.AddSection();
                section1.AddParagraph("Test Middle1 Section - 1");
                section1.AddHeadersAndFooters();
                section1.Header.Default.AddParagraph().AddText("Section 1 - Header");
                section1.Footer.Default.AddParagraph().AddText("Section 1 - Footer");

                var section2 = document.AddSection();
                section2.AddParagraph("Test Middle2 Section - 1");
                section2.AddHeadersAndFooters();
                section2.Header.Default.AddParagraph().AddText("Section 2 - Header");
                section2.Footer.Default.AddParagraph().AddText("Section 2 - Footer");

                var section3 = document.AddSection();
                section3.AddParagraph("Test Last Section - 1");
                section3.AddHeadersAndFooters();
                section3.DifferentOddAndEvenPages = true;
                section3.DifferentFirstPage = true;
                section3.Header.Default.AddParagraph().AddText("Section 3 - Header Odd/Default");
                section3.Footer.Default.AddParagraph().AddText("Section 3 - Footer Odd/Default");
                section3.Header.Even.AddParagraph().AddText("Section 3 - Header Even");
                section3.Footer.Even.AddParagraph().AddText("Section 3 - Footer Even");

                document.AddPageBreak();
                section3.AddParagraph("Test Last Section - 2");
                document.AddPageBreak();
                section3.AddParagraph("Test Last Section - 3");

                document.Save(openWord);

                Console.WriteLine("IsValid: " + document.DocumentIsValid);

                Console.WriteLine("Section 0 DifferentOddAndEventPages: " + document.Sections[0].DifferentOddAndEvenPages);
                Console.WriteLine("Section 0 DifferentFirstPage: " + document.Sections[0].DifferentFirstPage);

                Console.WriteLine("Section 1 DifferentOddAndEventPages: " + document.Sections[1].DifferentOddAndEvenPages);
                Console.WriteLine("Section 1 DifferentFirstPage: " + document.Sections[1].DifferentFirstPage);

                Console.WriteLine("Section 2 DifferentOddAndEventPages: " + document.Sections[2].DifferentOddAndEvenPages);
                Console.WriteLine("Section 2 DifferentFirstPage: " + document.Sections[2].DifferentFirstPage);

                Console.WriteLine("Section 3 DifferentOddAndEventPages: " + document.Sections[3].DifferentOddAndEvenPages);
                Console.WriteLine("Section 3 DifferentFirstPage: " + document.Sections[3].DifferentFirstPage);

            }
        }
    }
}
