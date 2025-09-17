using System;
using OfficeIMO.Examples.Utils;
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

            var headers = Guard.NotNull(document.Header, "Document headers must exist after enabling headers.");
            var defaultHeader = Guard.NotNull(headers.Default, "Default header must exist after enabling headers.");

            var headerParagraph = defaultHeader.AddParagraph("Header ");
            headerParagraph.AddText("clutter ");
            headerParagraph.AddText("text");
            defaultHeader.AddParagraph();

            var footers = Guard.NotNull(document.Footer, "Document footers must exist after enabling headers.");
            var defaultFooter = Guard.NotNull(footers.Default, "Default footer must exist after enabling headers.");

            var footerParagraph = defaultFooter.AddParagraph("Footer ");
            footerParagraph.AddText("clutter ");
            footerParagraph.AddText("text");
            defaultFooter.AddParagraph();

            document.CleanupDocument();
            document.Save(openWord);
        }
    }
}
