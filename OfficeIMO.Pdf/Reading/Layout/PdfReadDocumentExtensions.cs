using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>
/// Convenience APIs for column-aware text extraction at the document level.
/// </summary>
public static class PdfReadDocumentExtensions {
    /// <summary>
    /// Extracts text for all pages using simple two-column detection per page, separating pages with a blank line.
    /// </summary>
    /// <param name="options">Optional layout options controlling column detection, margins and trimming.</param>
    /// <returns>Concatenated plain text for all pages with inferred reading order.</returns>
    public static string ExtractTextWithColumns(this PdfReadDocument doc, PdfTextLayoutOptions? options = null) {
        var sb = new StringBuilder();
        for (int i = 0; i < doc.Pages.Count; i++) {
            if (i > 0) sb.AppendLine();
            sb.Append(doc.Pages[i].ExtractTextWithColumns(options));
        }
        return sb.ToString();
    }
}
