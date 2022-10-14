using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicWordWithDefaultStyleChange(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with different default style");
            string filePath = System.IO.Path.Combine(folderPath, "BasicWordWithDefaultStyleChange.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Settings.DefaultFontSize = 30;
                
                var paragraph1 = document.AddParagraph("Adding paragraph1 with some text and pressing ENTER");

                document.AddParagraph("Adding paragraph1 with some text and pressing ENTER").FontSize = 15;

                document.Save(openWord);
            }
        }
    }
}
