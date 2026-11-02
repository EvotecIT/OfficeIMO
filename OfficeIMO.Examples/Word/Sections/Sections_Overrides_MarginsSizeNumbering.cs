using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal partial class Sections
    {
        internal static void Example_Word_Fluent_Sections_Overrides_MarginsSizeNumbering(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with multiple sections and per-section overrides");
            string filePath = Path.Combine(folderPath, "Fluent_Sections_Overrides_MarginsSizeNumbering.docx");
            using (var document = WordDocument.Create(filePath))
            {
                document.PageOrientation = OfficePageOrientation.Portrait;
                document.PageSettings.PageSize = WordPageSize.A4;
                document.Sections[0].SetMargins(WordMargin.Normal);
                document.DifferentFirstPage = true;
                document.DifferentOddAndEvenPages = true;
                var first = document.AddSection(WordSectionBreakType.NextPage);
                first.SetMargins(WordMargin.Narrow);
                first.PageSettings.PageSize = WordPageSize.Legal;
                first.ColumnCount = 2;
                first.AddPageNumbering(1, WordNumberFormat.LowerRoman);
                first.AddParagraph("Section 1");
                first.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("Cell 1");
                var second = document.AddSection(WordSectionBreakType.NextPage);
                second.SetMargins(WordMargin.Wide);
                second.PageSettings.PageSize = WordPageSize.A3;
                second.AddPageNumbering(1);
                second.AddParagraph("Section 2");
                second.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("Cell 2");
                document.Save();
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}