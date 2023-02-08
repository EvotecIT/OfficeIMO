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

        public static void Example_EmbedFileMultiple(string folderPath, string templateFolder, bool openWord) {
            Console.WriteLine("[*] Creating standard document with multiple files embedded");
            string filePath = System.IO.Path.Combine(folderPath, "MultipleFilesEmedded.docx");

            string htmlFilePath = System.IO.Path.Combine(templateFolder, "SampleFileHTML.html");
            string rtfFilePath = System.IO.Path.Combine(templateFolder, "SampleFileRTF.rtf");
            string txtFilePath = System.IO.Path.Combine(templateFolder, "SampleFileTEXT.txt");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Add RTF document in front of the document");
                // we natively support HTML by using extension to decide what it is
                document.AddEmbeddedDocument(rtfFilePath);

                document.AddPageBreak();

                document.AddParagraph("Add HTML");

                // we natively support HTML by using extension to decide what it is
                document.AddEmbeddedDocument(htmlFilePath);

                document.AddPageBreak();

                document.AddParagraph("Add TEXT 1");

                // we natively support TextPlain by using extension (log/txt) to apply text/plain content type
                document.AddEmbeddedDocument(txtFilePath);

                document.AddPageBreak();

                document.AddParagraph("Add TEXT 2");

                // but you can specify type as you want, if the extension would be .ext, and the content type would be text/plain
                document.AddEmbeddedDocument(txtFilePath, AlternativeFormatImportPartType.TextPlain);


                Console.WriteLine("Embedded documents in word: " + document.EmbeddedDocuments.Count);
                Console.WriteLine("Embedded documents in Section 0: " + document.Sections[0].EmbeddedDocuments.Count);
                Console.WriteLine("Content type 0: " + document.EmbeddedDocuments[0].ContentType);
                Console.WriteLine("Content type 1: " + document.EmbeddedDocuments[1].ContentType);

                document.AddPageBreak();

                document.AddParagraph("Add RTF 2");

                document.AddEmbeddedDocument(rtfFilePath);

                Console.WriteLine("Content type 0: " + document.EmbeddedDocuments[0].ContentType);
                Console.WriteLine("Content type 1: " + document.EmbeddedDocuments[1].ContentType);
                Console.WriteLine("Content type 2: " + document.EmbeddedDocuments[2].ContentType);
                Console.WriteLine("Content type 3: " + document.EmbeddedDocuments[3].ContentType);
                Console.WriteLine("Content type 4: " + document.EmbeddedDocuments[4].ContentType);

                Console.WriteLine("Embedded documents in word: " + document.EmbeddedDocuments.Count);

                document.Save(openWord);
            }
        }
    }
}
