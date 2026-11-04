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
                document.PageSetup.Orientation = OfficePageOrientation.Portrait;
                document.PageSetup.Size.PageSize = WordPageSize.A4;
                document.Sections[0].PageSetup.Margins.Type = WordMargin.Normal;
                document.PageSetup.DifferentFirstPage = true;
                document.PageSetup.DifferentOddAndEvenPages = true;
                var first = document.AddSection(WordSectionBreakType.NextPage);
                first.PageSetup.Margins.Type = WordMargin.Narrow;
                first.PageSetup.Size.PageSize = WordPageSize.Legal;
                first.ColumnCount = 2;
                first.AddPageNumbering(1, WordNumberFormat.LowerRoman);
                first.AddParagraph("Section 1");
                first.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("Cell 1");
                var second = document.AddSection(WordSectionBreakType.NextPage);
                second.PageSetup.Margins.Type = WordMargin.Wide;
                second.PageSetup.Size.PageSize = WordPageSize.A3;
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