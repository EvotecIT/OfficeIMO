using System;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Examples.Utils;
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
                var section0 = document.Sections[0];
                var section0DefaultHeader = Guard.NotNull(section0.Header?.Default, "Section 0 should expose a default header after adding headers and footers.");
                section0DefaultHeader.AddParagraph().AddText("Section 0").AddBookmark("BookmarkInSection0Header1");

                var tableHeader = section0DefaultHeader.AddTable(3, 4);
                tableHeader.Rows[0].Cells[3].Paragraphs[0].Text = "This is sparta";
                Console.WriteLine(section0DefaultHeader.Tables.Count);

                section0DefaultHeader.AddHorizontalLine();

                section0DefaultHeader.AddHyperLink("Link to website!", new Uri("https://evotec.xyz"));

                section0DefaultHeader.AddHyperLink("Przemysław Klys Email Me", new Uri("mailto:kontakt@evotec.pl?subject=Test Subject"));

                section0DefaultHeader.AddField(WordFieldType.Author, WordFieldFormat.FirstCap);


                var section0DefaultFooter = Guard.NotNull(section0.Footer?.Default, "Section 0 should expose a default footer after adding headers and footers.");
                section0DefaultFooter.AddParagraph().AddText("Section 0").AddBookmark("BookmarkInSection0Header2");

                var tableFooter = section0DefaultFooter.AddTable(2, 3);
                tableFooter.Rows[0].Cells[2].Paragraphs[0].Text = "This is not sparta";

                section0DefaultFooter.AddHorizontalLine();

                section0DefaultFooter.AddHyperLink("Link to website!", new Uri("https://evotec.xyz"));

                section0DefaultFooter.AddHyperLink("Przemysław Klys Email Me", new Uri("mailto:kontakt@evotec.pl?subject=Test Subject"));

                section0DefaultFooter.AddField(WordFieldType.Author, WordFieldFormat.FirstCap);


                var section1 = document.AddSection();
                section1.AddParagraph("Test Middle1 Section - 1");
                section1.AddHeadersAndFooters();
                var section1DefaultHeader = Guard.NotNull(section1.Header?.Default, "Section 1 should expose a default header after adding headers and footers.");
                section1DefaultHeader.AddParagraph().AddText("Section 1 - Header");
                var section1DefaultFooter = Guard.NotNull(section1.Footer?.Default, "Section 1 should expose a default footer after adding headers and footers.");
                section1DefaultFooter.AddParagraph().AddText("Section 1 - Footer");

                var section2 = document.AddSection();
                section2.AddParagraph("Test Middle2 Section - 1");
                section2.AddHeadersAndFooters();
                var section2DefaultHeader = Guard.NotNull(section2.Header?.Default, "Section 2 should expose a default header after adding headers and footers.");
                section2DefaultHeader.AddParagraph().AddText("Section 2 - Header");
                var section2DefaultFooter = Guard.NotNull(section2.Footer?.Default, "Section 2 should expose a default footer after adding headers and footers.");
                section2DefaultFooter.AddParagraph().AddText("Section 2 - Footer");

                var section3 = document.AddSection();
                section3.AddParagraph("Test Last Section - 1");
                section3.AddHeadersAndFooters();
                section3.DifferentOddAndEvenPages = true;
                section3.DifferentFirstPage = true;
                var section3DefaultHeader = Guard.NotNull(section3.Header?.Default, "Section 3 should expose a default header after adding headers and footers.");
                section3DefaultHeader.AddParagraph().AddText("Section 3 - Header Odd/Default");
                var section3DefaultFooter = Guard.NotNull(section3.Footer?.Default, "Section 3 should expose a default footer after adding headers and footers.");
                section3DefaultFooter.AddParagraph().AddText("Section 3 - Footer Odd/Default");
                var section3EvenHeader = Guard.NotNull(section3.Header?.Even, "Section 3 should expose an even header after enabling different odd and even pages.");
                section3EvenHeader.AddParagraph().AddText("Section 3 - Header Even");
                var section3EvenFooter = Guard.NotNull(section3.Footer?.Even, "Section 3 should expose an even footer after enabling different odd and even pages.");
                section3EvenFooter.AddParagraph().AddText("Section 3 - Footer Even");

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
