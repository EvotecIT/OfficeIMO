using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class HyperLinks {

        internal static void Example_BasicWordWithHyperLinks(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with hyperlinks");
            string filePath = System.IO.Path.Combine(folderPath, "BasicDocumentHyperlinks.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test 1");

                document.AddParagraph("Hello users! Please visit ").AddHyperLink("bookmark below", "TestBookmark", true, "This is link to bookmark below shown within Tooltip");
                Console.WriteLine(document.HyperLinks.Count);
                Console.WriteLine(document.Sections[0].ParagraphsHyperLinks.Count);
                Console.WriteLine(document.ParagraphsHyperLinks.Count);
                Console.WriteLine(document.Sections[0].HyperLinks.Count);
                document.AddParagraph("Test HYPERLINK ").AddHyperLink(" to website?", new Uri("https://evotec.xyz"));

                document.AddParagraph("Test Email Address ").AddHyperLink("Przemysław Klys", new Uri("mailto:kontakt@evotec.pl?subject=Test Subject"));

                document.AddPageBreak();
                document.AddPageBreak();

                document.AddParagraph("Test HYPERLINK ").AddHyperLink(" to website?", new Uri("https://evotec.xyz"));
                document.AddParagraph("Test HYPERLINK ").AddHyperLink(" to website?", new Uri("https://evotec.pl"));

                //document.HyperLinks.Last().Remove();

                document.AddParagraph("Test 2").AddBookmark("TestBookmark");


                document.AddParagraph("Hello users! Please visit ").AddHyperLink("bookmark below", "TestBookmark", true, "This is link to bookmark below shown within Tooltip");


                document.HyperLinks.Last().Uri = new Uri("https://evotec.pl");
                document.HyperLinks.Last().Anchor = "";

                Console.WriteLine(document.HyperLinks.Count);

                document.Save(openWord);
            }
        }

    }
}
