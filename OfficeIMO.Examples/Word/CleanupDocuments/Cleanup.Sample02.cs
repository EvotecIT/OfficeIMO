using System;
using System.IO;
using OfficeIMO.Word;

internal static partial class CleanupDocuments {
    public static void CleanupDocuments_Sample02(string folderPath, bool openWord) {
        string filePath = System.IO.Path.Combine(folderPath, "SimpleWordDocumentReadyToCleanup1.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {

            document.AddParagraph("This is a text ").AddText("more text").AddText(" even longer text").AddText(" and even longer right?");

            Console.WriteLine("Paragraph count before merge: " + document.Paragraphs.Count);

            // since WordParagraph above are actually "Runs" with the same formatting cleanup will merge them as a single WordParagraph (single Run)
            var changesCount = document.CleanupDocument();
            Console.WriteLine("Changes count: " + changesCount);

            Console.WriteLine("Paragraph count after merging: " + document.Paragraphs.Count);

            Console.WriteLine("Merged text: " + document.Paragraphs[0].Text);

            document.Save(false);
        }
        using (WordDocument document = WordDocument.Load(Path.Combine(folderPath, "SimpleWordDocumentReadyToCleanup1.docx"))) {
            Console.WriteLine("Paragraph count after loading: " + document.Paragraphs.Count);

            Console.WriteLine("Merged text: " + document.Paragraphs[0].Text);

            document.AddParagraph("This is a text 1 ").AddText("more text 1").AddText(" even longer text 1").AddText(" and even longer right?");

            document.Paragraphs[3].Bold = true;
            document.Paragraphs[4].Bold = true;

            Console.WriteLine("Paragraph count before merge: " + document.Paragraphs.Count);

            document.CleanupDocument();

            Console.WriteLine("Paragraph count after merging: " + document.Paragraphs.Count);

            document.Save(false);
        }
    }
}
