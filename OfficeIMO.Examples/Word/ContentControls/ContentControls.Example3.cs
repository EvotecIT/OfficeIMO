using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class ContentControls {
        internal static void Example_AdvancedContentControls(string folderPath, bool openWord) {
            Console.WriteLine("[*] Advanced content control demo");
            string filePath = Path.Combine(folderPath, "DocumentAdvancedContentControls.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var para1 = document.AddParagraph("Control 1:");
                para1.AddStructuredDocumentTag(alias: "Alias1", text: "First", tag: "Tag1");

                var para2 = document.AddParagraph("Control 2:");
                para2.AddStructuredDocumentTag(alias: "Alias2", text: "Second", tag: "Tag2");

                var para3 = document.AddParagraph("Control 3:");
                para3.AddStructuredDocumentTag(alias: "Alias3", text: "Third", tag: "Tag3");

                document.Save(openWord);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var aliasControl = document.GetStructuredDocumentTagByAlias("Alias2");
                aliasControl.Text = "Changed";
                var tagControl = document.GetStructuredDocumentTagByTag("Tag3");
                Console.WriteLine("Tag3 text before: " + tagControl.Text);
                tagControl.Text = "Modified";
                document.Save(openWord);
            }
        }
    }
}
