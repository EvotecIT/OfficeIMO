using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System;
using System.IO;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_SaveShapes(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with shapes and exporting to PDF");
            string docPath = Path.Combine(folderPath, "ExportShapesToPdf.docx");
            string pdfPath = Path.Combine(folderPath, "ExportShapesToPdf.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddShape(ShapeType.Rectangle, 80, 40, Color.Aqua, Color.Black, 1);
                WordShape.AddLine(paragraph, 0, 50, 80, 50, Color.Red, 2);
                document.Save();
                document.SaveAsPdf(pdfPath);
            }
        }
    }
}
