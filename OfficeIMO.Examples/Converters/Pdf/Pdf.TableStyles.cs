using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System;
using System.IO;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_TableStyles(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating PDF with styled table");
            string docPath = Path.Combine(folderPath, "StyledTable.docx");
            string pdfPath = Path.Combine(folderPath, "StyledTable.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
                WordTableCell cell = table.Rows[0].Cells[0];
                cell.Paragraphs[0].Text = "Styled";
                cell.ShadingFillColor = OfficeColor.Parse("FFFF00");
                cell.Borders.TopStyle = WordBorderStyle.Single;
                cell.Borders.BottomStyle = WordBorderStyle.Single;
                cell.Borders.LeftStyle = WordBorderStyle.Single;
                cell.Borders.RightStyle = WordBorderStyle.Single;
                cell.Borders.TopColor = OfficeColor.Parse("FF0000");
                cell.Borders.BottomColor = OfficeColor.Parse("FF0000");
                cell.Borders.LeftColor = OfficeColor.Parse("FF0000");
                cell.Borders.RightColor = OfficeColor.Parse("FF0000");
                cell.Borders.TopSize = OfficeLength.Points(1);
                cell.Borders.BottomSize = OfficeLength.Points(1);
                cell.Borders.LeftSize = OfficeLength.Points(1);
                cell.Borders.RightSize = OfficeLength.Points(1);
                document.Save();
                document.SaveAsPdf(pdfPath);
            }
        }
    }
}
