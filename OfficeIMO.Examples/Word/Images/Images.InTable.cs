using System;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_AddingImagesSampleToTable(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with some Images and Samples");
            var filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithImagesSample4.docx");
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            using var document = WordDocument.Create(filePath);

            var table = document.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200);

            // not really nessessary to add new paragraph since one is already there by default
            var paragraph = table.Rows[0].Cells[1].AddParagraph();
            paragraph.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200);

            document.AddHeadersAndFooters();

            var tableInHeader = document.Header.Default.AddTable(2, 2);
            tableInHeader.Rows[0].Cells[0].Paragraphs[0].AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200);

            // not really nessessary to add new paragraph since one is already there by default
            var paragraphInHeader = tableInHeader.Rows[0].Cells[1].AddParagraph();
            paragraphInHeader.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 200, 200);

            document.Save(openWord);
        }
    }
}
