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

        public static void Example_EmbedFileHTML(string folderPath, string templateFolder, bool openWord) {
            Console.WriteLine("[*] Creating standard document with embedded HTML file");
            string filePath = System.IO.Path.Combine(folderPath, "EmbeddedFileHTML.docx");
            string htmlFilePath = System.IO.Path.Combine(templateFolder, "SampleFileHTML.html");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Add HTML document in DOCX");

                document.AddEmbeddedDocument(htmlFilePath);

                document.Save(openWord);
            }
        }
    }
}
