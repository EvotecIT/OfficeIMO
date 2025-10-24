using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word;

internal static partial class Paragraphs {

    internal static void Example_ParagraphParentNavigation(string folderPath, bool openWord) {
        Console.WriteLine("[*] Navigating from a paragraph to its owning cell");
        string filePath = Path.Combine(folderPath, "Paragraph Parent Navigation.docx");

        using (WordDocument document = WordDocument.Create(filePath)) {
            WordTable table = document.AddTable(2, 2, WordTableStyle.TableGrid);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Top left";

            var paragraph = table.Rows[1].Cells[1].AddParagraph("Parent lookup example");

            if (paragraph.Parent is WordTableCell parentCell) {
                var parentRow = parentCell.Parent;
                var parentTable = parentCell.ParentTable;

                Console.WriteLine($"Paragraph is stored in a cell that currently has {parentCell.Paragraphs.Count} paragraph(s).");
                Console.WriteLine($"The parent row exposes {parentRow.CellsCount} cell(s) and belongs to a table with {parentTable.RowsCount} row(s).");
                Console.WriteLine($"Table style: {parentTable.Style}");
            }

            document.Save(openWord);
        }
    }
}
