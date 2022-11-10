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
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;
                document.DifferentFirstPage = true;

                document.Header.Default.AddParagraph().SetColor(Color.Red).SetText("Test Header");

                document.Footer.Default.AddParagraph().SetColor(Color.Blue).SetText("Test Footer");

                Console.WriteLine("Header Default Count: " + document.Header.Default.Paragraphs.Count);
                Console.WriteLine("Header Even Count: " + document.Header.Even.Paragraphs.Count);
                Console.WriteLine("Header First Count: " + document.Header.First.Paragraphs.Count);

                Console.WriteLine("Header text: " + document.Header.Default.Paragraphs[0].Text);

                Console.WriteLine("Footer Default Count: " + document.Footer.Default.Paragraphs.Count);
                Console.WriteLine("Footer Even Count: " + document.Footer.Even.Paragraphs.Count);
                Console.WriteLine("Footer First Count: " + document.Footer.First.Paragraphs.Count);

                Console.WriteLine("Footer text: " + document.Footer.Default.Paragraphs[0].Text);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine("Header Default Count: " + document.Header.Default.Paragraphs.Count);
                Console.WriteLine("Header Even Count: " + document.Header.Even.Paragraphs.Count);
                Console.WriteLine("Header First Count: " + document.Header.First.Paragraphs.Count);

                Console.WriteLine("Header text: " + document.Header.Default.Paragraphs[0].Text);

                Console.WriteLine("Footer Default Count: " + document.Footer.Default.Paragraphs.Count);
                Console.WriteLine("Footer Even Count: " + document.Footer.Even.Paragraphs.Count);
                Console.WriteLine("Footer First Count: " + document.Footer.First.Paragraphs.Count);

                Console.WriteLine("Footer text: " + document.Footer.Default.Paragraphs[0].Text);

                document.Save(openWord);
            }
        }

    }
}
