using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class CrossReferences {
        internal static void Example_BasicCrossReferences(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with cross references");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithCrossReferences.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var heading = document.AddParagraph("Introduction");
                heading.Style = WordParagraphStyles.Heading1;
                heading.AddBookmark("Intro");

                document.AddParagraph("See chapter: ").AddCrossReference("Intro", WordCrossReferenceType.Heading).AddText(" for more information.");

                document.UpdateFields();
                document.Save(openWord);
            }
        }
    }
}
