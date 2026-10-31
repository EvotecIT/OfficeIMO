using System;
using System.IO;
using OfficeIMO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        public static void Example_FluentTableBuilder(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with fluent tables");
            string filePath = Path.Combine(folderPath, "FluentTableBuilder.docx");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    var table1 = document.AddTable(1, 3);
                    var row2 = table1.Rows[0];
                    row2.Cells[0].Paragraphs[0].Text = System.Convert.ToString("Name") ?? string.Empty;
                    row2.Cells[1].Paragraphs[0].Text = System.Convert.ToString("Role") ?? string.Empty;
                    row2.Cells[2].Paragraphs[0].Text = System.Convert.ToString("Score") ?? string.Empty;
                    row2.RepeatHeaderRowAtTheTopOfEachPage = true;
                    table1.ConditionalFormattingFirstRow = true;
                    var row3 = table1.AddRow();
                    row3.Cells[0].Paragraphs[0].Text = System.Convert.ToString("Alice") ?? string.Empty;
                    row3.Cells[1].Paragraphs[0].Text = System.Convert.ToString("Dev") ?? string.Empty;
                    row3.Cells[2].Paragraphs[0].Text = System.Convert.ToString(98) ?? string.Empty;
                    var row4 = table1.AddRow();
                    row4.Cells[0].Paragraphs[0].Text = System.Convert.ToString("Bob") ?? string.Empty;
                    row4.Cells[1].Paragraphs[0].Text = System.Convert.ToString("Ops") ?? string.Empty;
                    row4.Cells[2].Paragraphs[0].Text = System.Convert.ToString(91) ?? string.Empty;
                    table1.Style = WordTableStyle.TableGrid;
                    table1.Alignment = WordTableAlignment.Center;
                    table1.Width = WordTableWidth.Percent(100);
                    var table5 = document.AddTable(new object[, ] { { "Q", "Revenue", "Churn" }, { "Q1", "1.1M", "2.1%" }, { "Q2", "1.3M", "1.8%" } });
                    table5.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
                    var table6 = document.AddTable(2, 3);
                    for (int row = 0; row < table6.Rows.Count; row++)
                        for (int column = 0; column < table6.Rows[row].Cells.Count; column++)
                            (table6.Rows[row].Cells[column]).AddParagraph($"R{(row + 1)}C{(column + 1)}", true);
                    (table6.Rows[1 - 1].Cells[3 - 1]).AddParagraph("Last", true);
                    table6.InsertRow(3 - 1, "A", "B", "C");
                    table6.InsertColumn(4 - 1, "X", "Y", "Z");
                    (table6.Rows[1 - 1]).Cells.ForEach(c => c.ShadingFillColorHex = "ffcccc");
                    foreach (var row in table6.Rows)
                    {
                        (row.Cells[2 - 1]).ShadingFillColorHex = "ccffcc";
                    }

                    document.Save();
                }
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}