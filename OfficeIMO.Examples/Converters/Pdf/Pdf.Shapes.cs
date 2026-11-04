using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System;
using System.IO;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_SaveShapes(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with shapes and exporting to PDF");
            string docPath = Path.Combine(folderPath, "ExportShapesToPdf.docx");
            string pdfPath = Path.Combine(folderPath, "ExportShapesToPdf.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddShape(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(80), OfficeIMO.OfficeLength.Points(40), Color.Aqua, Color.Black, OfficeIMO.OfficeLength.Points(1));
                WordShape.AddLine(paragraph, OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(50), OfficeIMO.OfficeLength.Points(80), OfficeIMO.OfficeLength.Points(50), Color.Red, OfficeIMO.OfficeLength.Points(2));
                document.Save();
                document.SaveAsPdf(pdfPath);
            }
        }
    }
}
