using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal static partial class HyperLinks {

        internal static void Example_BasicWordWithHyperLinksInTables(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with hyperlinks");
            string filePath = System.IO.Path.Combine(folderPath, "BasicDocumentHyperlinksInTables.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test 1");

                document.AddParagraph("Hello users! Please visit ").AddHyperLink("bookmark below", "TestBookmark", true, "This is link to bookmark below shown within Tooltip");
                Console.WriteLine(document.HyperLinks.Count);
                Console.WriteLine(document.Sections[0].ParagraphsHyperLinks.Count);
                Console.WriteLine(document.ParagraphsHyperLinks.Count);
                Console.WriteLine(document.Sections[0].HyperLinks.Count);
                document.AddParagraph("Test HYPERLINK ").AddHyperLink(" to website?", new Uri("https://evotec.xyz"), true);

                document.AddTable(3, 3);

                document.Tables[0].Rows[0].Cells[0].Paragraphs[0].AddHyperLink(" to website?", new Uri("https://evotec.xyz"), addStyle: true);

                Console.WriteLine(document.Tables[0].Rows[0].Cells[0].Paragraphs[1].IsHyperLink);

                Console.WriteLine(document.HyperLinks.Count);

                document.Save(openWord);
            }
        }
    }
}
