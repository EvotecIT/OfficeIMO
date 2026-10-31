using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal partial class Sections
    {
        internal static void Example_Word_Fluent_Sections_PageSetup_Defaults(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with document-wide page setup defaults via fluent API");
            string filePath = Path.Combine(folderPath, "Fluent_Sections_PageSetup_Defaults.docx");
            using (var document = WordDocument.Create(filePath))
            {
                {
                    document.PageOrientation = OfficePageOrientation.Landscape;
                    document.PageSettings.PageSize = WordPageSize.A4;
                    document.Sections[0].SetMargins(WordMargin.Normal);
                    document.DifferentFirstPage = true;
                    document.DifferentOddAndEvenPages = true;
                    document.Save();
                }
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}