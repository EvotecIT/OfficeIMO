using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        internal static void Example_FluentTableCustomWidthAndShading(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Fluent table with custom widths and shading");
            string filePath = Path.Combine(folderPath, "FluentTableCustomWidthAndShading.docx");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    var table1 = document.AddTable(1, 2);
                    var row2 = table1.Rows[0];
                    row2.Cells[0].Paragraphs[0].Text = System.Convert.ToString("Red") ?? string.Empty;
                    row2.Cells[1].Paragraphs[0].Text = System.Convert.ToString("Blue") ?? string.Empty;
                    foreach (var row in table1.Rows)
                    {
                        row.Cells[1 - 1].WidthType = WordTableWidthUnit.Dxa;
                        row.Cells[1 - 1].Width = (int)(72 * 20);
                    }

                    foreach (var row in table1.Rows)
                    {
                        row.Cells[2 - 1].WidthType = WordTableWidthUnit.Dxa;
                        row.Cells[2 - 1].Width = (int)(144 * 20);
                    }

                    table1.Rows[1 - 1].Height = global::OfficeIMO.OfficeLength.Points(((int)(36 * 20)) / 20D);
                    {
                        (table1.Rows[1 - 1].Cells[1 - 1]).Paragraphs[0].ParagraphAlignment = WordParagraphAlignment.Center;
                        (table1.Rows[1 - 1].Cells[1 - 1]).ShadingFillColorHex = "ff0000";
                        (table1.Rows[1 - 1].Cells[1 - 1]).Borders.LeftStyle = WordBorderStyle.Single;
                        (table1.Rows[1 - 1].Cells[1 - 1]).Borders.RightStyle = WordBorderStyle.Single;
                        (table1.Rows[1 - 1].Cells[1 - 1]).Borders.TopStyle = WordBorderStyle.Single;
                        (table1.Rows[1 - 1].Cells[1 - 1]).Borders.BottomStyle = WordBorderStyle.Single;
                    }

                    {
                        (table1.Rows[1 - 1].Cells[2 - 1]).Paragraphs[0].ParagraphAlignment = WordParagraphAlignment.Center;
                        (table1.Rows[1 - 1].Cells[2 - 1]).ShadingFillColorHex = "0000ff";
                        (table1.Rows[1 - 1].Cells[2 - 1]).Borders.LeftStyle = WordBorderStyle.Single;
                        (table1.Rows[1 - 1].Cells[2 - 1]).Borders.RightStyle = WordBorderStyle.Single;
                        (table1.Rows[1 - 1].Cells[2 - 1]).Borders.TopStyle = WordBorderStyle.Single;
                        (table1.Rows[1 - 1].Cells[2 - 1]).Borders.BottomStyle = WordBorderStyle.Single;
                    }

                    document.Save();
                }
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}