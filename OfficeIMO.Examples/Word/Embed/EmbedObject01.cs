using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Embed {

        public static void Example_EmbedFileExcel(string folderPath, string templateFolder, bool openWord) {
            Console.WriteLine("[*] Creating standard document with embedded object (excel file)");
            string filePath = System.IO.Path.Combine(folderPath, "EmbeddedObjectExcel.docx");


            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Add Excel document in front of the document");

                string excelFilePath = System.IO.Path.Combine(templateFolder, "SampleFileExcel.xlsx");
                var imageFilePath = System.IO.Path.Combine(templateFolder, "SampleExcelIcon.png");

                document.AddEmbeddedObject(excelFilePath, imageFilePath);

                document.Save(openWord);
            }
        }
    }
}
