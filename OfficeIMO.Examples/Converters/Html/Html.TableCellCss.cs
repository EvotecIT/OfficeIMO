using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlTableCellCss(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlTableCellCss.docx");
            using var doc = WordDocument.Create();
            var table = doc.AddTable(1, 1);
            var cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = "Styled";
            cell.Paragraphs[0].ParagraphAlignment = JustificationValues.Right;
            cell.ShadingFillColorHex = "ff0000";
            cell.Borders.LeftStyle = BorderValues.Single;
            cell.Borders.RightStyle = BorderValues.Single;
            cell.Borders.TopStyle = BorderValues.Single;
            cell.Borders.BottomStyle = BorderValues.Single;
            cell.Borders.LeftColorHex = "00ff00";
            cell.Borders.RightColorHex = "00ff00";
            cell.Borders.TopColorHex = "00ff00";
            cell.Borders.BottomColorHex = "00ff00";
            cell.Borders.LeftSize = 8;
            cell.Borders.RightSize = 8;
            cell.Borders.TopSize = 8;
            cell.Borders.BottomSize = 8;
            doc.Save(filePath);

            string html = doc.ToHtml();
            Console.WriteLine(html);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
