using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicWordWithMargins(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with margins");
            string filePath = System.IO.Path.Combine(folderPath, "EmptyDocumentWithMargins.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].Margins.Bottom = 10;
                document.Sections[0].Margins.Left = 600;
                document.Sections[0].Margins.Top = 10;
                document.Sections[0].Margins.Right = 600;

                document.Settings.FontFamily = "Arial";
                document.Settings.FontSize = 9;

                document.AddParagraph("This should be Arial 9");

                var par = document.AddParagraph("This should be Tahoma 20");
                par.FontFamily = "Tahoma";
                par.FontSize = 20;

                document.Save(openWord);
            }
        }
    }
}
