using System;
using System.IO;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class ContentControls {
        internal static void Example_AdvancedContentControls(string folderPath, bool openWord) {
            Console.WriteLine("[*] Advanced content control demo");
            string filePath = Path.Combine(folderPath, "DocumentAdvancedContentControls.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var para1 = document.AddParagraph("Control 1:");
                para1.AddStructuredDocumentTag(text: "First", alias: "Alias1", tag: "Tag1");

                var para2 = document.AddParagraph("Control 2:");
                para2.AddStructuredDocumentTag(text: "Second", alias: "Alias2", tag: "Tag2");

                var para3 = document.AddParagraph("Control 3:");
                para3.AddStructuredDocumentTag(text: "Third", alias: "Alias3", tag: "Tag3");

                document.Save(openWord);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var aliasControl = Guard.NotNull(document.GetStructuredDocumentTagByAlias("Alias2"), "Content control with alias 'Alias2' was not found.");
                aliasControl.Text = "Changed";
                var tagControl = Guard.NotNull(document.GetStructuredDocumentTagByTag("Tag3"), "Content control with tag 'Tag3' was not found.");
                Console.WriteLine("Tag3 text before: " + tagControl.Text);
                tagControl.Text = "Modified";
                document.Save(openWord);
            }
        }
    }
}
