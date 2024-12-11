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
            string filePath = System.IO.Path.Combine(folderPath, "EmbeddedFileHTML3.docx");
            string htmlFilePath = System.IO.Path.Combine(templateFolder, "SampleFileHTML.html");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Console.WriteLine("Embedded documents in word: " + document.EmbeddedDocuments.Count);
                Console.WriteLine("Embedded documents in Section 0: " + document.Sections[0].EmbeddedDocuments.Count);

                document.AddParagraph("Add HTML document in DOCX");

                Console.WriteLine("Embedded documents in word: " + document.EmbeddedDocuments.Count);

                document.AddEmbeddedDocument(htmlFilePath);

                document.AddEmbeddedDocument(htmlFilePath);

                Console.WriteLine("Embedded documents in word: " + document.EmbeddedDocuments.Count);

                document.AddSection();

                document.AddEmbeddedDocument(htmlFilePath);

                document.AddEmbeddedDocument(htmlFilePath);

                Console.WriteLine("Embedded documents in word: " + document.EmbeddedDocuments.Count);
                Console.WriteLine("Embedded documents in Section 0: " + document.Sections[0].EmbeddedDocuments.Count);
                Console.WriteLine("Embedded documents in Section 1: " + document.Sections[1].EmbeddedDocuments.Count);

                document.AddEmbeddedDocument(htmlFilePath);

                Console.WriteLine("Embedded documents in word: " + document.EmbeddedDocuments.Count);
                Console.WriteLine("Embedded documents in Section 0: " + document.Sections[0].EmbeddedDocuments.Count);
                Console.WriteLine("Embedded documents in Section 1: " + document.Sections[1].EmbeddedDocuments.Count);

                document.EmbeddedDocuments[0].Save("C:\\TEMP\\EmbeddedFileHTML.html");

                document.AddEmbeddedDocument(htmlFilePath);

                Console.WriteLine("Embedded documents in word: " + document.EmbeddedDocuments.Count);
                Console.WriteLine("Embedded documents in Section 0: " + document.Sections[0].EmbeddedDocuments.Count);
                Console.WriteLine("Embedded documents in Section 1: " + document.Sections[1].EmbeddedDocuments.Count);
                Console.WriteLine("Content type: " + document.EmbeddedDocuments[0].ContentType);

                document.Save(openWord);
            }
        }
    }
}
