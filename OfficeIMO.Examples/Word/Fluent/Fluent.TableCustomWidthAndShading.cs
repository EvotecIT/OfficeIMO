using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        internal static void Example_FluentTableCustomWidthAndShading(string folderPath, bool openWord) {
            Console.WriteLine("[*] Fluent table with custom widths and shading");
            string filePath = Path.Combine(folderPath, "FluentTableCustomWidthAndShading.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Table(t => t
                        .Columns(2)
                        .Row("Red", "Blue")
                        .ColumnWidth(1, 72)
                        .ColumnWidth(2, 144)
                        .RowHeight(1, 36)
                        .CellStyle(1, 1, cell => {
                            cell.Paragraphs[0].ParagraphAlignment = JustificationValues.Center;
                            cell.ShadingFillColorHex = "ff0000";
                            cell.Borders.LeftStyle = BorderValues.Single;
                            cell.Borders.RightStyle = BorderValues.Single;
                            cell.Borders.TopStyle = BorderValues.Single;
                            cell.Borders.BottomStyle = BorderValues.Single;
                        })
                        .CellStyle(1, 2, cell => {
                            cell.Paragraphs[0].ParagraphAlignment = JustificationValues.Center;
                            cell.ShadingFillColorHex = "0000ff";
                            cell.Borders.LeftStyle = BorderValues.Single;
                            cell.Borders.RightStyle = BorderValues.Single;
                            cell.Borders.TopStyle = BorderValues.Single;
                            cell.Borders.BottomStyle = BorderValues.Single;
                        }))
                    .End()
                    .Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}

