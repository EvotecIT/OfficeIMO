using System;
using OfficeIMO.Word;

/// <summary>
/// Demonstrates cleanup operations on a newly created document.
/// </summary>
internal static partial class CleanupDocuments {
    /// <summary>
    /// Creates a document with redundant runs and applies cleanup.
    /// </summary>
    /// <param name="folderPath">Directory to create the file in.</param>
    /// <param name="openWord">Opens Word when <c>true</c>.</param>
    public static void CleanupDocuments_Sample03(string folderPath, bool openWord) {
        string filePath = System.IO.Path.Combine(folderPath, "CleanupDocument.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            WordParagraph p = document.AddParagraph("Hello");
            p.SetBold();
            p.AddText(" World").SetBold();
            document.CleanupDocument();
            document.Save(openWord);
        }
    }
}
