using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static partial class PdfTextExtractor {
    /// <summary>Extracts plain text from all pages, concatenated with blank lines between pages.</summary>
    public static string ExtractAllText(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        var bytes = File.ReadAllBytes(path);
        return ExtractAllText(bytes);
    }
    
    /// <summary>Extracts plain text from all pages using layout options such as column detection and header/footer trimming.</summary>
    public static string ExtractAllText(string path, PdfTextLayoutOptions? options) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        if (options is null) {
            return ExtractAllText(path);
        }
    
        return PdfReadDocument.Open(path).ExtractTextWithColumns(options);
    }
    
    /// <summary>Extracts plain text from all pages and writes UTF-8 text to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllText(string inputPath, Stream outputStream) {
        ExtractAllText(inputPath, outputStream, null);
    }
    
    /// <summary>Extracts plain text from all pages using layout options and writes UTF-8 text to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllText(string inputPath, Stream outputStream, PdfTextLayoutOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);
    
        WriteTextOutput(outputStream, ExtractAllText(inputPath, options));
    }
    
    /// <summary>Extracts plain text from all pages and writes UTF-8 text to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllText(string inputPath, string outputPath) {
        ExtractAllText(inputPath, outputPath, null);
    }
    
    /// <summary>Extracts plain text from all pages using layout options and writes UTF-8 text to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllText(string inputPath, string outputPath, PdfTextLayoutOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
    
        WriteTextOutput(fullOutputPath, ExtractAllText(inputPath, options));
    }
    
    /// <summary>Extracts plain text from each page in document order.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(string path) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractTextByPage(PdfReadDocument.Open(path));
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(string path, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractTextByPageRanges(PdfReadDocument.Open(path), pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges and concatenates selected pages with blank lines.</summary>
    public static string ExtractAllTextByPageRanges(string path, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractAllTextByPageRanges(PdfReadDocument.Open(path), null, pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges with layout options and concatenates selected pages with blank lines.</summary>
    public static string ExtractAllTextByPageRanges(string path, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractAllTextByPageRanges(PdfReadDocument.Open(path), options, pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges and writes one UTF-8 text result to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllTextByPageRanges(string inputPath, Stream outputStream, params PdfPageRange[] pageRanges) {
        ExtractAllTextByPageRanges(inputPath, outputStream, null, pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges with layout options and writes one UTF-8 text result to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllTextByPageRanges(string inputPath, Stream outputStream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);
    
        WriteTextOutput(outputStream, ExtractAllTextByPageRanges(inputPath, options, pageRanges));
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges and writes one UTF-8 text result to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllTextByPageRanges(string inputPath, string outputPath, params PdfPageRange[] pageRanges) {
        ExtractAllTextByPageRanges(inputPath, outputPath, null, pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges with layout options and writes one UTF-8 text result to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllTextByPageRanges(string inputPath, string outputPath, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
    
        WriteTextOutput(fullOutputPath, ExtractAllTextByPageRanges(inputPath, options, pageRanges));
    }
    
    /// <summary>Extracts logical Markdown from all pages.</summary>
    public static string ExtractMarkdown(string path, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return PdfLogicalDocument.Load(path, options).ToMarkdown(markdownOptions);
    }
    
    /// <summary>Extracts logical Markdown from all pages and writes UTF-8 Markdown to <paramref name="outputStream"/>.</summary>
    public static void ExtractMarkdown(string inputPath, Stream outputStream, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);
        WriteTextOutput(outputStream, ExtractMarkdown(inputPath, options, markdownOptions));
    }
    
    /// <summary>Extracts logical Markdown from all pages and writes UTF-8 Markdown to <paramref name="outputPath"/>.</summary>
    public static void ExtractMarkdown(string inputPath, string outputPath, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteTextOutput(fullOutputPath, ExtractMarkdown(inputPath, options, markdownOptions));
    }
    
    /// <summary>Extracts logical Markdown from each page in document order.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPage(string path, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractMarkdownByPage(PdfLogicalDocument.Load(path, options), markdownOptions);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPageRanges(string path, params PdfPageRange[] pageRanges) {
        return ExtractMarkdownByPageRanges(path, null, null, pageRanges);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPageRanges(string path, PdfTextLayoutOptions? options, PdfLogicalMarkdownOptions? markdownOptions, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractMarkdownByPage(PdfLogicalDocument.LoadPageRanges(path, options, pageRanges), markdownOptions);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges and concatenates selected pages with Markdown page separators.</summary>
    public static string ExtractMarkdownByPageRangesAsDocument(string path, params PdfPageRange[] pageRanges) {
        return ExtractMarkdownByPageRangesAsDocument(path, null, null, pageRanges);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges and concatenates selected pages with Markdown page separators.</summary>
    public static string ExtractMarkdownByPageRangesAsDocument(string path, PdfTextLayoutOptions? options, PdfLogicalMarkdownOptions? markdownOptions, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return PdfLogicalDocument.LoadPageRanges(path, options, pageRanges).ToMarkdown(markdownOptions);
    }
    
    /// <summary>Extracts structured content for each page, including detected lines, lists, leader rows, and simple tables.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPage(string path, PdfTextLayoutOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return PdfReadDocument.Open(path).ExtractStructuredPages(options);
    }
    
    /// <summary>Extracts structured content from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPageRanges(string path, params PdfPageRange[] pageRanges) {
        return ExtractStructuredByPageRanges(path, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts structured content from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPageRanges(string path, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractStructuredByPageRanges(PdfReadDocument.Open(path), options, pageRanges);
    }
    
    /// <summary>Extracts detected paragraphs grouped by page while preserving paragraph geometry.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPage(string path, PdfTextLayoutOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return PdfReadDocument.Open(path).ExtractParagraphsByPage(options);
    }
    
    /// <summary>Extracts detected paragraphs from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPageRanges(string path, params PdfPageRange[] pageRanges) {
        return ExtractParagraphsByPageRanges(path, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected paragraphs from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPageRanges(string path, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractParagraphsByPageRanges(PdfReadDocument.Open(path), options, pageRanges);
    }
    
    /// <summary>Extracts detected headings grouped by page while preserving heading geometry.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPage(string path, PdfTextLayoutOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return PdfReadDocument.Open(path).ExtractHeadingsByPage(options);
    }
    
    /// <summary>Extracts detected headings from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPageRanges(string path, params PdfPageRange[] pageRanges) {
        return ExtractHeadingsByPageRanges(path, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected headings from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPageRanges(string path, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractHeadingsByPageRanges(PdfReadDocument.Open(path), options, pageRanges);
    }
    
    /// <summary>Extracts detected list items grouped by page while preserving marker and nesting hints.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPage(string path, PdfTextLayoutOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return PdfReadDocument.Open(path).ExtractListItemsByPage(options);
    }
    
    /// <summary>Extracts detected list items from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPageRanges(string path, params PdfPageRange[] pageRanges) {
        return ExtractListItemsByPageRanges(path, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected list items from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPageRanges(string path, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractListItemsByPageRanges(PdfReadDocument.Open(path), options, pageRanges);
    }
    
    /// <summary>Extracts detected tables grouped by page while preserving table geometry.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPage(string path, PdfTextLayoutOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return PdfReadDocument.Open(path).ExtractTablesByPage(options);
    }
    
    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPageRanges(string path, params PdfPageRange[] pageRanges) {
        return ExtractTablesByPageRanges(path, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges in caller order.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPageRanges(string path, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return ExtractTablesByPageRanges(PdfReadDocument.Open(path), options, pageRanges);
    }
    
    /// <summary>Extracts detected tables and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPage(string inputPath, string outputDirectory, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var tablePages = ExtractTablesByPage(inputPath, options);
        return WriteTableCsvFiles(inputPath, fullOutputDirectory, tablePages);
    }
    
    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPageRanges(string inputPath, string outputDirectory, params PdfPageRange[] pageRanges) {
        return ExtractTablesByPageRanges(inputPath, outputDirectory, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPageRanges(string inputPath, string outputDirectory, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var tablePages = ExtractTablesByPageRanges(inputPath, options, pageRanges);
        return WriteTableCsvFiles(inputPath, fullOutputDirectory, tablePages);
    }
    
    /// <summary>Extracts detected tables from the current stream position and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPage(Stream stream, string outputDirectory, string baseName = "table", PdfTextLayoutOptions? options = null) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var tablePages = ExtractTablesByPage(stream, options);
        return WriteTableCsvFiles(baseName, fullOutputDirectory, tablePages);
    }
    
    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges from the current stream position and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPageRanges(Stream stream, string outputDirectory, string baseName = "table", params PdfPageRange[] pageRanges) {
        return ExtractTablesByPageRanges(stream, outputDirectory, baseName, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges from the current stream position and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPageRanges(Stream stream, string outputDirectory, string baseName, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var tablePages = ExtractTablesByPageRanges(stream, options, pageRanges);
        return WriteTableCsvFiles(baseName, fullOutputDirectory, tablePages);
    }
    
    /// <summary>Extracts detected tables from bytes and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPage(byte[] pdf, string outputDirectory, string baseName = "table", PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var tablePages = ExtractTablesByPage(pdf, options);
        return WriteTableCsvFiles(baseName, fullOutputDirectory, tablePages);
    }
    
    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges from bytes and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPageRanges(byte[] pdf, string outputDirectory, string baseName = "table", params PdfPageRange[] pageRanges) {
        return ExtractTablesByPageRanges(pdf, outputDirectory, baseName, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges from bytes and writes one CSV file per detected table.</summary>
    public static IReadOnlyList<string> ExtractTablesByPageRanges(byte[] pdf, string outputDirectory, string baseName, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var tablePages = ExtractTablesByPageRanges(pdf, options, pageRanges);
        return WriteTableCsvFiles(baseName, fullOutputDirectory, tablePages);
    }
    
    /// <summary>Extracts plain text from each page using layout options such as column detection and header/footer trimming.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(string path, PdfTextLayoutOptions? options) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        if (options is null) {
            return ExtractTextByPage(path);
        }
    
        return ExtractTextByPage(PdfReadDocument.Open(path), options);
    }
    
    /// <summary>Extracts plain text from each page and writes one UTF-8 text file per page.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(string inputPath, string outputDirectory) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractTextByPage(inputPath);
        return WriteTextPages(inputPath, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges and writes one UTF-8 text file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(string inputPath, string outputDirectory, params PdfPageRange[] pageRanges) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedTextPages(PdfReadDocument.Open(inputPath), pageRanges);
        return WriteTextPages(inputPath, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges with layout options and writes one UTF-8 text file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(string inputPath, string outputDirectory, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedTextPages(PdfReadDocument.Open(inputPath), options, pageRanges);
        return WriteTextPages(inputPath, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts plain text from each page with layout options and writes one UTF-8 text file per page.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(string inputPath, string outputDirectory, PdfTextLayoutOptions? options) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractTextByPage(inputPath, options);
        return WriteTextPages(inputPath, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts logical Markdown from each page and writes one UTF-8 Markdown file per page.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPage(string inputPath, string outputDirectory, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractMarkdownByPage(inputPath, options, markdownOptions);
        return WriteMarkdownPages(inputPath, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges and writes one UTF-8 Markdown file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPageRanges(string inputPath, string outputDirectory, params PdfPageRange[] pageRanges) {
        return ExtractMarkdownByPageRanges(inputPath, outputDirectory, null, null, pageRanges);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges and writes one UTF-8 Markdown file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPageRanges(string inputPath, string outputDirectory, PdfTextLayoutOptions? options, PdfLogicalMarkdownOptions? markdownOptions, params PdfPageRange[] pageRanges) {
        Guard.NotNull(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedMarkdownPages(PdfLogicalDocument.LoadPageRanges(inputPath, options, pageRanges), markdownOptions);
        return WriteMarkdownPages(inputPath, fullOutputDirectory, pages);
    }
}
