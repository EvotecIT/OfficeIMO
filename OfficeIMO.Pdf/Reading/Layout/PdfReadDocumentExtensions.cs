using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>
/// Convenience APIs for column-aware text extraction at the document level.
/// </summary>
public static class PdfReadDocumentExtensions {
    /// <summary>
    /// Extracts text for all pages using simple two-column detection per page, separating pages with a blank line.
    /// </summary>
    /// <param name="doc">Source document.</param>
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

    /// <summary>
    /// Extracts simple structured content (lines, TOC rows, list items, leader rows) for the whole document.
    /// </summary>
    /// <param name="doc">Source document.</param>
    /// <param name="options">Optional layout options.</param>
    public static (List<string> Lines, List<(string Label, int Page)> Toc, List<string> Lists, List<string[]> LeaderRows)
        ExtractStructured(this PdfReadDocument doc, PdfTextLayoutOptions? options = null) {
        var lines = new List<string>();
        var toc = new List<(string, int)>();
        var lists = new List<string>();
        var leaders = new List<string[]>();
        for (int i = 0; i < doc.Pages.Count; i++) {
            var s = doc.Pages[i].ExtractStructured(options);
            lines.AddRange(s.Lines);
            toc.AddRange(s.Toc);
            lists.AddRange(s.ListItems);
            leaders.AddRange(s.LeaderRows);
        }
        return (lines, toc, lists, leaders);
    }
}
