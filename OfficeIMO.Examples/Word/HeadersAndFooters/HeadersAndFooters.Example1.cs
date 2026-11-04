using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Word {
    internal static partial class HeadersAndFooters {
        internal static void Example_BasicWordWithHeaderAndFooter0(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Headers and Footers");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Headers and Footers Default.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.PageSetup.DifferentOddAndEvenPages = true;
                document.PageSetup.DifferentFirstPage = true;

                var setup = document.PageSetup;
                var defaultHeader = setup.Header;
                var evenHeader = setup.EvenPageHeader;
                var firstHeader = setup.FirstPageHeader;

                defaultHeader.AddParagraph().SetColor(Color.Red).SetText("Test Header");

                var defaultFooter = setup.Footer;
                var evenFooter = setup.EvenPageFooter;
                var firstFooter = setup.FirstPageFooter;

                defaultFooter.AddParagraph().SetColor(Color.Blue).SetText("Test Footer");

                Console.WriteLine("Header Default Count: " + defaultHeader.Paragraphs.Count);
                Console.WriteLine("Header Even Count: " + evenHeader.Paragraphs.Count);
                Console.WriteLine("Header First Count: " + firstHeader.Paragraphs.Count);

                Console.WriteLine("Header text: " + defaultHeader.Paragraphs[0].Text);

                Console.WriteLine("Footer Default Count: " + defaultFooter.Paragraphs.Count);
                Console.WriteLine("Footer Even Count: " + evenFooter.Paragraphs.Count);
                Console.WriteLine("Footer First Count: " + firstFooter.Paragraphs.Count);

                Console.WriteLine("Footer text: " + defaultFooter.Paragraphs[0].Text);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var setup = document.PageSetup;
                var defaultHeader = setup.Header;
                var evenHeader = setup.EvenPageHeader;
                var firstHeader = setup.FirstPageHeader;

                Console.WriteLine("Header Default Count: " + defaultHeader.Paragraphs.Count);
                Console.WriteLine("Header Even Count: " + evenHeader.Paragraphs.Count);
                Console.WriteLine("Header First Count: " + firstHeader.Paragraphs.Count);

                Console.WriteLine("Header text: " + defaultHeader.Paragraphs[0].Text);

                var defaultFooter = setup.Footer;
                var evenFooter = setup.EvenPageFooter;
                var firstFooter = setup.FirstPageFooter;

                Console.WriteLine("Footer Default Count: " + defaultFooter.Paragraphs.Count);
                Console.WriteLine("Footer Even Count: " + evenFooter.Paragraphs.Count);
                Console.WriteLine("Footer First Count: " + firstFooter.Paragraphs.Count);

                Console.WriteLine("Footer text: " + defaultFooter.Paragraphs[0].Text);

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }

    }
}
