using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class HeadersAndFooters {
        internal static void Example_BasicWordWithHeaderAndFooter0(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Headers and Footers");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Headers and Footers Default.docx");

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
                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;
                document.DifferentFirstPage = true;

                var headers = RequireHeaders(document.Header, "Document headers");
                var defaultHeader = RequireHeader(headers.Default, "Default header");
                var evenHeader = RequireHeader(headers.Even, "Even header");
                var firstHeader = RequireHeader(headers.First, "First header");

                defaultHeader.AddParagraph().SetColor(Color.Red).SetText("Test Header");

                var footers = RequireFooters(document.Footer, "Document footers");
                var defaultFooter = RequireFooter(footers.Default, "Default footer");
                var evenFooter = RequireFooter(footers.Even, "Even footer");
                var firstFooter = RequireFooter(footers.First, "First footer");

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
                var headers = RequireHeaders(document.Header, "Document headers");
                var defaultHeader = RequireHeader(headers.Default, "Default header");
                var evenHeader = RequireHeader(headers.Even, "Even header");
                var firstHeader = RequireHeader(headers.First, "First header");

                Console.WriteLine("Header Default Count: " + defaultHeader.Paragraphs.Count);
                Console.WriteLine("Header Even Count: " + evenHeader.Paragraphs.Count);
                Console.WriteLine("Header First Count: " + firstHeader.Paragraphs.Count);

                Console.WriteLine("Header text: " + defaultHeader.Paragraphs[0].Text);

                var footers = RequireFooters(document.Footer, "Document footers");
                var defaultFooter = RequireFooter(footers.Default, "Default footer");
                var evenFooter = RequireFooter(footers.Even, "Even footer");
                var firstFooter = RequireFooter(footers.First, "First footer");

                Console.WriteLine("Footer Default Count: " + defaultFooter.Paragraphs.Count);
                Console.WriteLine("Footer Even Count: " + evenFooter.Paragraphs.Count);
                Console.WriteLine("Footer First Count: " + firstFooter.Paragraphs.Count);

                Console.WriteLine("Footer text: " + defaultFooter.Paragraphs[0].Text);

                document.Save(openWord);
            }
        }

    }
}
