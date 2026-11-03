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
                    document.PageSetup.Orientation = OfficePageOrientation.Landscape;
                    document.PageSetup.Size.PageSize = WordPageSize.A4;
                    document.Sections[0].PageSetup.Margins.Type = WordMargin.Normal;
                    document.PageSetup.DifferentFirstPage = true;
                    document.PageSetup.DifferentOddAndEvenPages = true;
                    document.Save();
                }
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}