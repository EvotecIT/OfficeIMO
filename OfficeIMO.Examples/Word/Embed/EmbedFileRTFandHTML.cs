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

        public static void Example_EmbedFileRTFandHTML(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with embedded RTF & HTML file");
            string filePath = System.IO.Path.Combine(folderPath, "EmbeddedFileRTFandHTML.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Add RTF document in front of the document");

                document.AddEmbeddedDocument(@"C:\Users\przemyslaw.klys\Downloads\file-sample_100kB.rtf");

                document.AddPageBreak();

                document.AddParagraph("Add HTML document as last in the document");

                document.AddEmbeddedDocument(@"C:\Users\przemyslaw.klys\Downloads\The global structure of an HTML document.html");

                document.Save(openWord);
            }
        }
    }
}
