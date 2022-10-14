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
                document.Settings.FontSize = 30;
                document.Settings.FontFamily = "Calibri Light";
                document.Settings.Language = "pl-PL";

                var paragraph1 = document.AddParagraph("To jest po polsku");

                var paragraph2 = document.AddParagraph("Adding paragraph1 with some text and pressing ENTER");
                paragraph2.FontSize = 15;
                paragraph2.FontFamily = "Courier New";

                document.Save(openWord);
            }
        }
    }
}
