using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BasicDocument {
        public static void Example_BasicWordWithPolishChars(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with polish chars");
            string filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithPolishChars.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Adding paragraph with some text with special chars to check if FontFamily works correctly for those");

                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                // this is only a test of setting FontFamily per paragraph. Please use document.Settings.FontFamily to set it per document.
                paragraph = document.AddParagraph("Wszedł kot do domu, gdzie były różne buty. ");
                paragraph.FontFamily = "Courier New";
                paragraph = paragraph.AddText("Chodził tak sobie i chodził, i się nachodził. ");
                paragraph.FontFamily = "Courier New";
                paragraph = paragraph.AddText("A potem jeszcze pochodził, i wąchał. A wszystko to nagrało życie. ");
                paragraph.FontFamily = "Courier New";

                paragraph = document.AddParagraph("English العربية");
                paragraph.FontFamily = "Courier New"; // overwrites all styles
                paragraph.FontFamilyEastAsia = "Arial"; // to change east asia font family
                paragraph.FontFamilyHighAnsi = "Arial"; // to change high ansi font family
                paragraph.FontFamilyComplexScript = "Arial"; // to change complex script font family
                document.Save(openWord);
            }
        }
    }
}
