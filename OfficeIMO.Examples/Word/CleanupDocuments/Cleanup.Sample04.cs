using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word;

/// <summary>
/// Demonstrates cleanup of headers and footers.
/// </summary>
internal static partial class CleanupDocuments {
    /// <summary>
    /// Creates a document with redundant runs and empty paragraphs in headers and footers and cleans it up.
    /// </summary>
    /// <param name="folderPath">Directory to create the file in.</param>
    /// <param name="openWord">Opens Word when <c>true</c>.</param>
    public static void CleanupDocuments_Sample04(string folderPath, bool openWord) {
        string filePath = System.IO.Path.Combine(folderPath, "CleanupHeadersFooters.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            document.AddHeadersAndFooters();

            var headerParagraph = document.Header!.Default.AddParagraph("Header ");
            headerParagraph.AddText("clutter ");
            headerParagraph.AddText("text");
            document.Header!.Default.AddParagraph();

            var footerParagraph = document.Footer!.Default.AddParagraph("Footer ");
            footerParagraph.AddText("clutter ");
            footerParagraph.AddText("text");
            document.Footer!.Default.AddParagraph();

            document.CleanupDocument();
            document.Save(openWord);
        }
    }
}
