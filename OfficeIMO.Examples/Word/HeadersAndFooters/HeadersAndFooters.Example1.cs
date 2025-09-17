using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
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

                var defaultHeader = GetDocumentHeaderOrThrow(document);
                defaultHeader.AddParagraph().SetColor(SixLabors.ImageSharp.Color.Red).SetText("Test Header");

                var defaultFooter = GetDocumentFooterOrThrow(document);
                defaultFooter.AddParagraph().SetColor(SixLabors.ImageSharp.Color.Blue).SetText("Test Footer");

                var evenHeader = GetDocumentHeaderOrThrow(document, HeaderFooterValues.Even);
                var firstHeader = GetDocumentHeaderOrThrow(document, HeaderFooterValues.First);

                Console.WriteLine("Header Default Count: " + defaultHeader.Paragraphs.Count);
                Console.WriteLine("Header Even Count: " + evenHeader.Paragraphs.Count);
                Console.WriteLine("Header First Count: " + firstHeader.Paragraphs.Count);

                Console.WriteLine("Header text: " + defaultHeader.Paragraphs[0].Text);

                var evenFooter = GetDocumentFooterOrThrow(document, HeaderFooterValues.Even);
                var firstFooter = GetDocumentFooterOrThrow(document, HeaderFooterValues.First);

                Console.WriteLine("Footer Default Count: " + defaultFooter.Paragraphs.Count);
                Console.WriteLine("Footer Even Count: " + evenFooter.Paragraphs.Count);
                Console.WriteLine("Footer First Count: " + firstFooter.Paragraphs.Count);

                Console.WriteLine("Footer text: " + defaultFooter.Paragraphs[0].Text);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var defaultHeader = GetDocumentHeaderOrThrow(document);
                var evenHeader = GetDocumentHeaderOrThrow(document, HeaderFooterValues.Even);
                var firstHeader = GetDocumentHeaderOrThrow(document, HeaderFooterValues.First);

                Console.WriteLine("Header Default Count: " + defaultHeader.Paragraphs.Count);
                Console.WriteLine("Header Even Count: " + evenHeader.Paragraphs.Count);
                Console.WriteLine("Header First Count: " + firstHeader.Paragraphs.Count);

                Console.WriteLine("Header text: " + defaultHeader.Paragraphs[0].Text);

                var defaultFooter = GetDocumentFooterOrThrow(document);
                var evenFooter = GetDocumentFooterOrThrow(document, HeaderFooterValues.Even);
                var firstFooter = GetDocumentFooterOrThrow(document, HeaderFooterValues.First);

                Console.WriteLine("Footer Default Count: " + defaultFooter.Paragraphs.Count);
                Console.WriteLine("Footer Even Count: " + evenFooter.Paragraphs.Count);
                Console.WriteLine("Footer First Count: " + firstFooter.Paragraphs.Count);

                Console.WriteLine("Footer text: " + defaultFooter.Paragraphs[0].Text);

                document.Save(openWord);
            }
        }

    }
}
