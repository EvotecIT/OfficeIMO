using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicWordWithTabs(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with tabs");
            string filePath = System.IO.Path.Combine(folderPath, "BasicWordWithTabs.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph1 = document.AddParagraph("To jest po polsku").AddTab().AddTab().AddText("Test");

                Console.WriteLine(document.Paragraphs.Count);

                Console.WriteLine(document.Paragraphs[1].IsTab);
                Console.WriteLine(document.Paragraphs[2].IsTab);

                Console.WriteLine(paragraph1.IsTab);

                var paragraph2 = document.AddParagraph("Adding paragraph1 with some text and pressing ENTER").AddTab();

                Console.WriteLine(document.Paragraphs.Count);

                Console.WriteLine(paragraph2.IsTab);

                paragraph2.Tab.Remove();
                document.Save(openWord);
            }
        }
    }
}
