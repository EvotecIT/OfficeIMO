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
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<StructuredPage> ExtractStructuredByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
    
        var pages = new List<StructuredPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            pages.Add(document.Pages[pageNumbers[i] - 1].ExtractStructured(options));
        }
    
        return pages.AsReadOnly();
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<StructuredParagraphPage> ExtractParagraphsByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
    
        var pages = new List<StructuredParagraphPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            var structuredPage = document.Pages[pageNumber - 1].ExtractStructured(options);
            pages.Add(new StructuredParagraphPage(pageNumber, structuredPage.Paragraphs));
        }
    
        return pages.AsReadOnly();
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<StructuredHeadingPage> ExtractHeadingsByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
    
        var pages = new List<StructuredHeadingPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            var structuredPage = document.Pages[pageNumber - 1].ExtractStructured(options);
            pages.Add(new StructuredHeadingPage(pageNumber, structuredPage.Headings));
        }
    
        return pages.AsReadOnly();
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<StructuredListItemPage> ExtractListItemsByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
    
        var pages = new List<StructuredListItemPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            var structuredPage = document.Pages[pageNumber - 1].ExtractStructured(options);
            pages.Add(new StructuredListItemPage(pageNumber, structuredPage.ListNodes));
        }
    
        return pages.AsReadOnly();
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<StructuredTablePage> ExtractTablesByPageRanges(PdfReadDocument document, PdfTextLayoutOptions? options, PdfPageRange[] pageRanges) {
        int[] pageNumbers = PdfPageRange.ExpandMany(pageRanges, document.Pages.Count, nameof(pageRanges));
    
        var pages = new List<StructuredTablePage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            var structuredPage = document.Pages[pageNumber - 1].ExtractStructured(options);
            pages.Add(new StructuredTablePage(pageNumber, structuredPage.TablesDetailed));
        }
    
        return pages.AsReadOnly();
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<string> ExtractMarkdownByPage(PdfLogicalDocument document, PdfLogicalMarkdownOptions? markdownOptions) {
        var pages = new List<string>(document.Pages.Count);
        for (int i = 0; i < document.Pages.Count; i++) {
            pages.Add(document.Pages[i].ToMarkdown(markdownOptions));
        }
    
        return pages.AsReadOnly();
    }
    
    private static System.Collections.ObjectModel.ReadOnlyCollection<SelectedTextPage> ExtractSelectedMarkdownPages(PdfLogicalDocument document, PdfLogicalMarkdownOptions? markdownOptions) {
        var pages = new List<SelectedTextPage>(document.Pages.Count);
        for (int i = 0; i < document.Pages.Count; i++) {
            PdfLogicalPage page = document.Pages[i];
            pages.Add(new SelectedTextPage(page.PageNumber, page.ToMarkdown(markdownOptions)));
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
