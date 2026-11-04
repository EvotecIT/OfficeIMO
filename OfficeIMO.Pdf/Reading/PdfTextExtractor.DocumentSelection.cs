using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static partial class PdfTextExtractor {
    internal static System.Collections.ObjectModel.ReadOnlyCollection<string> ExtractTextByPage(PdfReadDocument document) {
        var pages = new List<string>(document.Pages.Count);
        for (int i = 0; i < document.Pages.Count; i++) {
            pages.Add(document.Pages[i].ExtractText());
        }
    
        return pages.AsReadOnly();
    }
    
    internal static System.Collections.ObjectModel.ReadOnlyCollection<string> ExtractTextByPageRanges(PdfReadDocument document, PdfPageRange[] pageRanges) {
        var selected = ExtractSelectedTextPages(document, pageRanges);
        var pages = new List<string>(selected.Count);
        for (int i = 0; i < selected.Count; i++) {
            pages.Add(selected[i].Text);
        }
    
        return pages.AsReadOnly();
    }
    
    internal static string ExtractAllTextByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        var selected = ExtractSelectedTextPages(document, options, pageRanges);
        var sb = new StringBuilder();
        for (int i = 0; i < selected.Count; i++) {
            if (i > 0) {
                sb.AppendLine();
            }
    
            sb.Append(selected[i].Text);
        }
    
        return sb.ToString();
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<SelectedTextPage> ExtractSelectedTextPages(PdfReadDocument document, PdfPageRange[] pageRanges) {
        return ExtractSelectedTextPages(document, null, pageRanges);
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<SelectedTextPage> ExtractSelectedTextPages(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
    
        var pages = new List<SelectedTextPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            string text = options is null
                ? document.Pages[pageNumber - 1].ExtractText()
                : document.Pages[pageNumber - 1].ExtractTextWithColumns(options);
            pages.Add(new SelectedTextPage(pageNumber, text));
        }
    
        return pages.AsReadOnly();
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfStructuredPage> ExtractStructuredByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
    
        var pages = new List<PdfStructuredPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            pages.Add(document.Pages[pageNumbers[i] - 1].ExtractStructured(options));
        }
    
        return pages.AsReadOnly();
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfStructuredParagraphPage> ExtractParagraphsByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
    
        var pages = new List<PdfStructuredParagraphPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            var structuredPage = document.Pages[pageNumber - 1].ExtractStructured(options);
            pages.Add(new PdfStructuredParagraphPage(pageNumber, structuredPage.Paragraphs));
        }
    
        return pages.AsReadOnly();
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfStructuredHeadingPage> ExtractHeadingsByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
    
        var pages = new List<PdfStructuredHeadingPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            var structuredPage = document.Pages[pageNumber - 1].ExtractStructured(options);
            pages.Add(new PdfStructuredHeadingPage(pageNumber, structuredPage.Headings));
        }
    
        return pages.AsReadOnly();
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfStructuredListItemPage> ExtractListItemsByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
    
        var pages = new List<PdfStructuredListItemPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            var structuredPage = document.Pages[pageNumber - 1].ExtractStructured(options);
            pages.Add(new PdfStructuredListItemPage(pageNumber, structuredPage.ListNodes));
        }
    
        return pages.AsReadOnly();
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfStructuredTablePage> ExtractTablesByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
    
        var pages = new List<PdfStructuredTablePage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            var structuredPage = document.Pages[pageNumber - 1].ExtractStructured(options);
            pages.Add(new PdfStructuredTablePage(pageNumber, structuredPage.TablesDetailed));
        }
    
        return pages.AsReadOnly();
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<string> ExtractMarkdownByPage(PdfDocumentReadResult document, PdfLogicalMarkdownOptions? markdownOptions) {
        var pages = new List<string>(document.Pages.Count);
        for (int i = 0; i < document.Pages.Count; i++) {
            pages.Add(document.Pages[i].ExtractMarkdown(markdownOptions));
        }
    
        return pages.AsReadOnly();
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<SelectedTextPage> ExtractSelectedMarkdownPages(PdfDocumentReadResult document, PdfLogicalMarkdownOptions? markdownOptions) {
        var pages = new List<SelectedTextPage>(document.Pages.Count);
        for (int i = 0; i < document.Pages.Count; i++) {
            PdfLogicalPage page = document.Pages[i];
            pages.Add(new SelectedTextPage(page.PageNumber, page.ExtractMarkdown(markdownOptions)));
        }
    
        return pages.AsReadOnly();
    }
    
    private readonly struct SelectedTextPage {
        internal SelectedTextPage(int pageNumber, string text) {
            PageNumber = pageNumber;
            Text = text;
        }
    
        internal int PageNumber { get; }
        internal string Text { get; }
    }
    
    internal static System.Collections.ObjectModel.ReadOnlyCollection<string> ExtractTextByPage(PdfReadDocument document, PdfTextLayoutOptions options) {
        var pages = new List<string>(document.Pages.Count);
        for (int i = 0; i < document.Pages.Count; i++) {
            pages.Add(document.Pages[i].ExtractTextWithColumns(options));
        }
    
        return pages.AsReadOnly();
    }
}
