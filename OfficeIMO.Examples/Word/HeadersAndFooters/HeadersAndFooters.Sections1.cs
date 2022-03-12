using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class HeadersAndFooters {
        public static void Sections1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Sections - Headers/Footers");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with Sections - HeadersAndFooters.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph");

                var section1 = document.AddSection();
                section1.AddParagraph("Test Middle1 Section - 1");

                var section2 = document.AddSection();
                section2.AddParagraph("Test Middle2 Section - 1");

                var section3 = document.AddSection();
                section3.AddParagraph("Test Last Section - 1");

                document.Save(openWord);
            }
        }
    }
}
