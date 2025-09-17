using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class HeadersAndFooters {
        internal static void Example_BasicWordWithHeaderAndFooter0(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Headers and Footers");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Headers and Footers Default.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;
                document.DifferentFirstPage = true;

                var defaultHeader = Guard.NotNull(document.Header?.Default, "Default header should exist after calling AddHeadersAndFooters.");
                defaultHeader.AddParagraph().SetColor(Color.Red).SetText("Test Header");

                var defaultFooter = Guard.NotNull(document.Footer?.Default, "Default footer should exist after calling AddHeadersAndFooters.");
                defaultFooter.AddParagraph().SetColor(Color.Blue).SetText("Test Footer");

                var evenHeader = Guard.NotNull(document.Header?.Even, "Even header should exist after enabling different odd and even pages.");
                var firstHeader = Guard.NotNull(document.Header?.First, "First header should exist after enabling different first page.");

                var evenFooter = Guard.NotNull(document.Footer?.Even, "Even footer should exist after enabling different odd and even pages.");
                var firstFooter = Guard.NotNull(document.Footer?.First, "First footer should exist after enabling different first page.");

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
                var loadDefaultHeader = Guard.NotNull(document.Header?.Default, "Default header should exist when reloading the document.");
                var loadEvenHeader = Guard.NotNull(document.Header?.Even, "Even header should exist when reloading the document.");
                var loadFirstHeader = Guard.NotNull(document.Header?.First, "First header should exist when reloading the document.");

                Console.WriteLine("Header Default Count: " + loadDefaultHeader.Paragraphs.Count);
                Console.WriteLine("Header Even Count: " + loadEvenHeader.Paragraphs.Count);
                Console.WriteLine("Header First Count: " + loadFirstHeader.Paragraphs.Count);

                Console.WriteLine("Header text: " + loadDefaultHeader.Paragraphs[0].Text);

                var loadDefaultFooter = Guard.NotNull(document.Footer?.Default, "Default footer should exist when reloading the document.");
                var loadEvenFooter = Guard.NotNull(document.Footer?.Even, "Even footer should exist when reloading the document.");
                var loadFirstFooter = Guard.NotNull(document.Footer?.First, "First footer should exist when reloading the document.");

                Console.WriteLine("Footer Default Count: " + loadDefaultFooter.Paragraphs.Count);
                Console.WriteLine("Footer Even Count: " + loadEvenFooter.Paragraphs.Count);
                Console.WriteLine("Footer First Count: " + loadFirstFooter.Paragraphs.Count);

                Console.WriteLine("Footer text: " + loadDefaultFooter.Paragraphs[0].Text);

                document.Save(openWord);
            }
        }

    }
}
