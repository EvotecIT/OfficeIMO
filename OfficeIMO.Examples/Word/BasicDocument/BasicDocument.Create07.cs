using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicWordWithSomeParagraphs(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with margins");
            string filePath = System.IO.Path.Combine(folderPath, "EmptyDocumentWithSomeParagraphs.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Settings.FontFamily = "Arial";
                document.Settings.FontSize = 9;
                document.AddParagraph("This should be Arial 9");

                var par = document.AddParagraph("This should be Tahoma 20");
                par.FontFamily = "Tahoma";
                par.AddText("SuperScript").SetVerticalTextAlignment(WordVerticalTextPosition.Superscript);
                par.AddText("Continue 1 ");
                par.AddText("Baseline").SetVerticalTextAlignment(WordVerticalTextPosition.Baseline);
                par.AddText("Continue 2 ");
                par.AddText("SubScript").SetVerticalTextAlignment(WordVerticalTextPosition.Subscript);

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
