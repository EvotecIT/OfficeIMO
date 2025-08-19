using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word.Sections {
    internal static class Sections_PageSetup_Defaults {
        public static void Example_Word_Fluent_Sections_PageSetup_Defaults(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with document-wide page setup defaults via fluent API");
            string filePath = Path.Combine(folderPath, "Fluent_Sections_PageSetup_Defaults.docx");

            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .PageSetup(ps => ps
                        .Orientation(PageOrientationValues.Landscape)
                        .Size(WordPageSize.A4)
                        .Margins(WordMargin.Normal)
                        .DifferentFirstPage()
                        .DifferentOddAndEvenPages())
                    .End();

                document.Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
