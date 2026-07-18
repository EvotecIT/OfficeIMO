using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static partial class PdfTextExtractor {
    /// <summary>Gets document metadata from the canonical parsed model.</summary>
    public static (string? Title, string? Author, string? Subject, string? Keywords) GetMetadata(
        byte[] pdf,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return GetMetadata(PdfReadDocument.Open(pdf, readOptions));
    }

    /// <summary>Gets document metadata from a bounded file snapshot.</summary>
    public static (string? Title, string? Author, string? Subject, string? Keywords) GetMetadata(
        string path,
        PdfReadOptions? readOptions = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return GetMetadata(PdfReadDocument.Open(path, readOptions));
    }

    /// <summary>
    /// Gets document metadata from a complete stream snapshot. Seekable streams are restored.
    /// </summary>
    public static (string? Title, string? Author, string? Subject, string? Keywords) GetMetadata(
        Stream stream,
        PdfReadOptions? readOptions = null) {
        Guard.NotNull(stream, nameof(stream));
        return GetMetadata(PdfReadDocument.Open(stream, readOptions));
    }

    private static (string? Title, string? Author, string? Subject, string? Keywords) GetMetadata(
        PdfReadDocument document) {
        PdfMetadata metadata = document.Metadata;
        return (metadata.Title, metadata.Author, metadata.Subject, metadata.Keywords);
    }

    /// <summary>Extracts plain text from all pages, concatenated with blank lines between pages.</summary>
    public static string ExtractAllText(byte[] pdf) {
        return ExtractAllText(pdf, (PdfTextLayoutOptions?)null, (PdfReadOptions?)null);
    }

    /// <summary>Extracts plain text from all pages, concatenated with blank lines between pages.</summary>
    public static string ExtractAllText(byte[] pdf, PdfTextLayoutOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        return options is null
            ? document.ExtractText()
            : document.ExtractTextWithColumns(options);
    }
    
    /// <summary>Extracts plain text from all pages using layout options such as column detection and header/footer trimming.</summary>
    public static string ExtractAllText(byte[] pdf, PdfTextLayoutOptions? options) {
        return ExtractAllText(pdf, options, null);
    }
    
    /// <summary>Extracts plain text from all pages and writes UTF-8 text to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllText(byte[] pdf, Stream outputStream) {
        ExtractAllText(pdf, outputStream, null);
    }
    
    /// <summary>Extracts plain text from all pages using layout options and writes UTF-8 text to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllText(byte[] pdf, Stream outputStream, PdfTextLayoutOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));
        ValidateWritableOutputStream(outputStream);
    
        WriteTextOutput(outputStream, ExtractAllText(pdf, options, null));
    }
    
    /// <summary>Extracts plain text from all pages and writes UTF-8 text to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllText(byte[] pdf, string outputPath) {
        ExtractAllText(pdf, outputPath, null);
    }
    
    /// <summary>Extracts plain text from all pages using layout options and writes UTF-8 text to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllText(byte[] pdf, string outputPath, PdfTextLayoutOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));
        string fullOutputPath = ValidateOutputPath(outputPath);
    
        WriteTextOutput(fullOutputPath, ExtractAllText(pdf, options, null));
    }
    
    /// <summary>Extracts plain text from each page in document order.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractTextByPage(PdfReadDocument.Open(pdf));
    }

    /// <summary>Extracts plain text from each page in document order.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(byte[] pdf, PdfTextLayoutOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        if (options is null) {
            return ExtractTextByPage(PdfReadDocument.Open(pdf, readOptions));
        }

        return ExtractTextByPage(PdfReadDocument.Open(pdf, readOptions), options);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractTextByPageRanges(PdfReadDocument.Open(pdf), pageRanges);
    }

    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(byte[] pdf, PdfPageRange[] pageRanges, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractTextByPageRanges(PdfReadDocument.Open(pdf, readOptions), pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges and concatenates selected pages with blank lines.</summary>
    public static string ExtractAllTextByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractAllTextByPageRanges(PdfReadDocument.Open(pdf), null, pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges with layout options and concatenates selected pages with blank lines.</summary>
    public static string ExtractAllTextByPageRanges(byte[] pdf, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractAllTextByPageRanges(PdfReadDocument.Open(pdf), options, pageRanges);
    }

    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges with layout options and concatenates selected pages with blank lines.</summary>
    public static string ExtractAllTextByPageRanges(byte[] pdf, PdfTextLayoutOptions? options, PdfReadOptions? readOptions, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractAllTextByPageRanges(PdfReadDocument.Open(pdf, readOptions), options, pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges and writes one UTF-8 text result to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllTextByPageRanges(byte[] pdf, Stream outputStream, params PdfPageRange[] pageRanges) {
        ExtractAllTextByPageRanges(pdf, outputStream, null, pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges with layout options and writes one UTF-8 text result to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllTextByPageRanges(byte[] pdf, Stream outputStream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        ValidateWritableOutputStream(outputStream);
    
        WriteTextOutput(outputStream, ExtractAllTextByPageRanges(pdf, options, pageRanges));
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges and writes one UTF-8 text result to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllTextByPageRanges(byte[] pdf, string outputPath, params PdfPageRange[] pageRanges) {
        ExtractAllTextByPageRanges(pdf, outputPath, null, pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges with layout options and writes one UTF-8 text result to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllTextByPageRanges(byte[] pdf, string outputPath, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        string fullOutputPath = ValidateOutputPath(outputPath);
    
        WriteTextOutput(fullOutputPath, ExtractAllTextByPageRanges(pdf, options, pageRanges));
    }
    
    /// <summary>Extracts logical Markdown from all pages.</summary>
    public static string ExtractMarkdown(byte[] pdf, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfLogicalDocument.Load(pdf, options).ToMarkdown(markdownOptions);
    }

    /// <summary>Extracts logical Markdown from all pages.</summary>
    public static string ExtractMarkdown(byte[] pdf, PdfTextLayoutOptions? options, PdfLogicalMarkdownOptions? markdownOptions, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfLogicalDocument.From(PdfReadDocument.Open(pdf, readOptions), options).ToMarkdown(markdownOptions);
    }
    
    /// <summary>Extracts logical Markdown from all pages and writes UTF-8 Markdown to <paramref name="outputStream"/>.</summary>
    public static void ExtractMarkdown(byte[] pdf, Stream outputStream, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        ValidateWritableOutputStream(outputStream);
        WriteTextOutput(outputStream, ExtractMarkdown(pdf, options, markdownOptions));
    }
    
    /// <summary>Extracts logical Markdown from all pages and writes UTF-8 Markdown to <paramref name="outputPath"/>.</summary>
    public static void ExtractMarkdown(byte[] pdf, string outputPath, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteTextOutput(fullOutputPath, ExtractMarkdown(pdf, options, markdownOptions));
    }
    
    /// <summary>Extracts logical Markdown from each page in document order.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPage(byte[] pdf, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractMarkdownByPage(PdfLogicalDocument.Load(pdf, options), markdownOptions);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ExtractMarkdownByPageRanges(pdf, (PdfTextLayoutOptions?)null, (PdfLogicalMarkdownOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPageRanges(byte[] pdf, PdfTextLayoutOptions? options, PdfLogicalMarkdownOptions? markdownOptions, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractMarkdownByPage(PdfLogicalDocument.LoadPageRanges(pdf, options, pageRanges), markdownOptions);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges and concatenates selected pages with Markdown page separators.</summary>
    public static string ExtractMarkdownByPageRangesAsDocument(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ExtractMarkdownByPageRangesAsDocument(pdf, (PdfTextLayoutOptions?)null, (PdfLogicalMarkdownOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges and concatenates selected pages with Markdown page separators.</summary>
    public static string ExtractMarkdownByPageRangesAsDocument(byte[] pdf, PdfTextLayoutOptions? options, PdfLogicalMarkdownOptions? markdownOptions, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfLogicalDocument.LoadPageRanges(pdf, options, pageRanges).ToMarkdown(markdownOptions);
    }

    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges and concatenates selected pages with Markdown page separators.</summary>
    public static string ExtractMarkdownByPageRangesAsDocument(byte[] pdf, PdfTextLayoutOptions? options, PdfLogicalMarkdownOptions? markdownOptions, PdfReadOptions? readOptions, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfLogicalDocument.FromPageRanges(PdfReadDocument.Open(pdf, readOptions), options, pageRanges).ToMarkdown(markdownOptions);
    }
    
    /// <summary>Extracts logical Markdown from each page from bytes and writes one UTF-8 Markdown file per page.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPage(byte[] pdf, string outputDirectory, string baseName = "page", PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractMarkdownByPage(pdf, options, markdownOptions);
        return WriteMarkdownPages(baseName, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges from bytes and writes one UTF-8 Markdown file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPageRanges(byte[] pdf, string outputDirectory, string baseName = "page", params PdfPageRange[] pageRanges) {
        return ExtractMarkdownByPageRanges(pdf, outputDirectory, baseName, null, null, pageRanges);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges from bytes and writes one UTF-8 Markdown file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPageRanges(byte[] pdf, string outputDirectory, string baseName, PdfTextLayoutOptions? options, PdfLogicalMarkdownOptions? markdownOptions, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedMarkdownPages(PdfLogicalDocument.LoadPageRanges(pdf, options, pageRanges), markdownOptions);
        return WriteMarkdownPages(baseName, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts structured content for each page, including detected lines, lists, leader rows, and simple tables.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPage(byte[] pdf, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfReadDocument.Open(pdf).ExtractStructuredPages(options);
    }
    
    /// <summary>Extracts structured content from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ExtractStructuredByPageRanges(pdf, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts structured content from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPageRanges(byte[] pdf, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractStructuredByPageRanges(PdfReadDocument.Open(pdf), options, pageRanges);
    }
    
    /// <summary>Extracts detected paragraphs grouped by page while preserving paragraph geometry.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPage(byte[] pdf, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfReadDocument.Open(pdf).ExtractParagraphsByPage(options);
    }
    
    /// <summary>Extracts detected paragraphs from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ExtractParagraphsByPageRanges(pdf, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected paragraphs from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPageRanges(byte[] pdf, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractParagraphsByPageRanges(PdfReadDocument.Open(pdf), options, pageRanges);
    }
    
    /// <summary>Extracts detected headings grouped by page while preserving heading geometry.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPage(byte[] pdf, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfReadDocument.Open(pdf).ExtractHeadingsByPage(options);
    }
    
    /// <summary>Extracts detected headings from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ExtractHeadingsByPageRanges(pdf, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected headings from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPageRanges(byte[] pdf, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractHeadingsByPageRanges(PdfReadDocument.Open(pdf), options, pageRanges);
    }
    
    /// <summary>Extracts detected list items grouped by page while preserving marker and nesting hints.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPage(byte[] pdf, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfReadDocument.Open(pdf).ExtractListItemsByPage(options);
    }
    
    /// <summary>Extracts detected list items from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ExtractListItemsByPageRanges(pdf, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected list items from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPageRanges(byte[] pdf, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractListItemsByPageRanges(PdfReadDocument.Open(pdf), options, pageRanges);
    }
    
    /// <summary>Extracts detected tables grouped by page while preserving table geometry.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPage(byte[] pdf, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return PdfReadDocument.Open(pdf).ExtractTablesByPage(options);
    }
    
    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ExtractTablesByPageRanges(pdf, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPageRanges(byte[] pdf, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        return ExtractTablesByPageRanges(PdfReadDocument.Open(pdf), options, pageRanges);
    }
    
    /// <summary>Extracts plain text from each page using layout options such as column detection and header/footer trimming.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(byte[] pdf, PdfTextLayoutOptions? options) {
        return ExtractTextByPage(pdf, options, null);
    }
}
