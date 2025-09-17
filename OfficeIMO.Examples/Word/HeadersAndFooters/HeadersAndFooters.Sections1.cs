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

            WordHeaders RequireHeaders(WordHeaders? headers, string description) {
                if (headers == null) {
                    throw new InvalidOperationException($"{description} are not available.");
                }

                return headers;
            }

            WordHeader RequireHeader(WordHeader? header, string description) {
                if (header == null) {
                    throw new InvalidOperationException($"{description} is not available.");
                }

                return header;
            }

            WordFooters RequireFooters(WordFooters? footers, string description) {
                if (footers == null) {
                    throw new InvalidOperationException($"{description} are not available.");
                }

                return footers;
            }

            WordFooter RequireFooter(WordFooter? footer, string description) {
                if (footer == null) {
                    throw new InvalidOperationException($"{description} is not available.");
                }

                return footer;
            }

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
                var section0Headers = RequireHeaders(section0.Header, "Section 0 headers");
                var section0DefaultHeader = RequireHeader(section0Headers.Default, "Section 0 default header");

                section0DefaultHeader.AddParagraph().AddText("Section 0").AddBookmark("BookmarkInSection0Header1");

                var tableHeader = section0DefaultHeader.AddTable(3, 4);
                tableHeader.Rows[0].Cells[3].Paragraphs[0].Text = "This is sparta";
                Console.WriteLine(section0DefaultHeader.Tables.Count);

                section0DefaultHeader.AddHorizontalLine();

                section0DefaultHeader.AddHyperLink("Link to website!", new Uri("https://evotec.xyz"));

                section0DefaultHeader.AddHyperLink("Przemysław Klys Email Me", new Uri("mailto:kontakt@evotec.pl?subject=Test Subject"));

                section0DefaultHeader.AddField(WordFieldType.Author, WordFieldFormat.FirstCap);


                var section0Footers = RequireFooters(section0.Footer, "Section 0 footers");
                var section0DefaultFooter = RequireFooter(section0Footers.Default, "Section 0 default footer");

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
                var section1Headers = RequireHeaders(section1.Header, "Section 1 headers");
                var section1DefaultHeader = RequireHeader(section1Headers.Default, "Section 1 default header");
                section1DefaultHeader.AddParagraph().AddText("Section 1 - Header");
                var section1Footers = RequireFooters(section1.Footer, "Section 1 footers");
                var section1DefaultFooter = RequireFooter(section1Footers.Default, "Section 1 default footer");
                section1DefaultFooter.AddParagraph().AddText("Section 1 - Footer");

                var section2 = document.AddSection();
                section2.AddParagraph("Test Middle2 Section - 1");
                section2.AddHeadersAndFooters();
                var section2Headers = RequireHeaders(section2.Header, "Section 2 headers");
                var section2DefaultHeader = RequireHeader(section2Headers.Default, "Section 2 default header");
                section2DefaultHeader.AddParagraph().AddText("Section 2 - Header");
                var section2Footers = RequireFooters(section2.Footer, "Section 2 footers");
                var section2DefaultFooter = RequireFooter(section2Footers.Default, "Section 2 default footer");
                section2DefaultFooter.AddParagraph().AddText("Section 2 - Footer");

                var section3 = document.AddSection();
                section3.AddParagraph("Test Last Section - 1");
                section3.AddHeadersAndFooters();
                section3.DifferentOddAndEvenPages = true;
                section3.DifferentFirstPage = true;
                var section3Headers = RequireHeaders(section3.Header, "Section 3 headers");
                var section3DefaultHeader = RequireHeader(section3Headers.Default, "Section 3 default header");
                section3DefaultHeader.AddParagraph().AddText("Section 3 - Header Odd/Default");
                var section3Footers = RequireFooters(section3.Footer, "Section 3 footers");
                var section3DefaultFooter = RequireFooter(section3Footers.Default, "Section 3 default footer");
                section3DefaultFooter.AddParagraph().AddText("Section 3 - Footer Odd/Default");
                var section3EvenHeader = RequireHeader(section3Headers.Even, "Section 3 even header");
                section3EvenHeader.AddParagraph().AddText("Section 3 - Header Even");
                var section3EvenFooter = RequireFooter(section3Footers.Even, "Section 3 even footer");
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
