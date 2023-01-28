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

        public static void Example_EmbedFileRTFandHTMLandTOC(string folderPath, string templateFolder, bool openWord) {
            Console.WriteLine("[*] Creating standard document with embedded RTF & HTML file");
            string filePath = System.IO.Path.Combine(folderPath, "EmbeddedFileRTFandHTMLandTOC.docx");

            string htmlFilePath = System.IO.Path.Combine(templateFolder, "SampleFileHTML.html");
            string rtfFilePath = System.IO.Path.Combine(templateFolder, "SampleFileRTF.rtf");

            using (WordDocument document = WordDocument.Create(filePath)) {

                document.AddTableOfContent();

                document.AddPageBreak();

                WordList wordList = document.AddList(WordListStyle.Headings111);
                wordList.AddItem("Embedded RTF").Style = WordParagraphStyles.Heading1;

                document.AddParagraph("Add RTF document in front of the document");

                document.AddEmbeddedDocument(rtfFilePath);

                document.AddPageBreak();

                wordList.AddItem("Embedded HTML").Style = WordParagraphStyles.Heading1;

                document.AddParagraph("Add HTML document as last in the document");

                document.AddEmbeddedDocument(htmlFilePath);

                document.Settings.UpdateFieldsOnOpen = true;
                document.Save(openWord);
            }
        }
    }
}
