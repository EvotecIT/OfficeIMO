using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static partial class PdfTextExtractor {
    /// <summary>Extracts plain text from all pages from the current stream position, concatenated with blank lines between pages.</summary>
    public static string ExtractAllText(Stream stream) {
        return ExtractAllText(ReadAllBytes(stream));
    }
    
    /// <summary>Extracts plain text from all pages from the current stream position using layout options such as column detection and header/footer trimming.</summary>
    public static string ExtractAllText(Stream stream, PdfTextLayoutOptions? options) {
        if (options is null) {
            return ExtractAllText(stream);
        }
    
        return PdfReadDocument.Open(stream).ExtractTextWithColumns(options);
    }
    
    /// <summary>Extracts plain text from all pages from the current position of a readable stream and writes UTF-8 text to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllText(Stream inputStream, Stream outputStream) {
        ExtractAllText(inputStream, outputStream, null);
    }
    
    /// <summary>Extracts plain text from all pages from the current position of a readable stream using layout options and writes UTF-8 text to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllText(Stream inputStream, Stream outputStream, PdfTextLayoutOptions? options) {
        ValidateWritableOutputStream(outputStream);
        WriteTextOutput(outputStream, ExtractAllText(inputStream, options));
    }
    
    /// <summary>Extracts plain text from all pages from the current position of a readable stream and writes UTF-8 text to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllText(Stream inputStream, string outputPath) {
        ExtractAllText(inputStream, outputPath, null);
    }
    
    /// <summary>Extracts plain text from all pages from the current position of a readable stream using layout options and writes UTF-8 text to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllText(Stream inputStream, string outputPath, PdfTextLayoutOptions? options) {
        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteTextOutput(fullOutputPath, ExtractAllText(inputStream, options));
    }
    
    /// <summary>Extracts plain text from each page from the current position of a readable stream.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(Stream stream) {
        return ExtractTextByPage(PdfReadDocument.Open(stream));
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from the current position of a readable stream.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractTextByPageRanges(PdfReadDocument.Open(stream), pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from the current position of a readable stream and concatenates selected pages with blank lines.</summary>
    public static string ExtractAllTextByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractAllTextByPageRanges(PdfReadDocument.Open(stream), null, pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from the current position of a readable stream with layout options and concatenates selected pages with blank lines.</summary>
    public static string ExtractAllTextByPageRanges(Stream stream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        return ExtractAllTextByPageRanges(PdfReadDocument.Open(stream), options, pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from the current position of a readable stream and writes one UTF-8 text result to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllTextByPageRanges(Stream inputStream, Stream outputStream, params PdfPageRange[] pageRanges) {
        ExtractAllTextByPageRanges(inputStream, outputStream, null, pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from the current position of a readable stream with layout options and writes one UTF-8 text result to <paramref name="outputStream"/>.</summary>
    public static void ExtractAllTextByPageRanges(Stream inputStream, Stream outputStream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        ValidateWritableOutputStream(outputStream);
        WriteTextOutput(outputStream, ExtractAllTextByPageRanges(inputStream, options, pageRanges));
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from the current position of a readable stream and writes one UTF-8 text result to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllTextByPageRanges(Stream inputStream, string outputPath, params PdfPageRange[] pageRanges) {
        ExtractAllTextByPageRanges(inputStream, outputPath, null, pageRanges);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from the current position of a readable stream with layout options and writes one UTF-8 text result to <paramref name="outputPath"/>.</summary>
    public static void ExtractAllTextByPageRanges(Stream inputStream, string outputPath, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteTextOutput(fullOutputPath, ExtractAllTextByPageRanges(inputStream, options, pageRanges));
    }
    
    /// <summary>Extracts logical Markdown from all pages from the current position of a readable stream.</summary>
    public static string ExtractMarkdown(Stream stream, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        return PdfLogicalDocument.Load(stream, options).ToMarkdown(markdownOptions);
    }
    
    /// <summary>Extracts logical Markdown from all pages from the current position of a readable stream and writes UTF-8 Markdown to <paramref name="outputStream"/>.</summary>
    public static void ExtractMarkdown(Stream inputStream, Stream outputStream, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        ValidateWritableOutputStream(outputStream);
        WriteTextOutput(outputStream, ExtractMarkdown(inputStream, options, markdownOptions));
    }
    
    /// <summary>Extracts logical Markdown from all pages from the current position of a readable stream and writes UTF-8 Markdown to <paramref name="outputPath"/>.</summary>
    public static void ExtractMarkdown(Stream inputStream, string outputPath, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        string fullOutputPath = ValidateOutputPath(outputPath);
        WriteTextOutput(fullOutputPath, ExtractMarkdown(inputStream, options, markdownOptions));
    }
    
    /// <summary>Extracts logical Markdown from each page from the current position of a readable stream.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPage(Stream stream, PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        return ExtractMarkdownByPage(PdfLogicalDocument.Load(stream, options), markdownOptions);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges from the current position of a readable stream.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractMarkdownByPageRanges(stream, (PdfTextLayoutOptions?)null, (PdfLogicalMarkdownOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges from the current position of a readable stream.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPageRanges(Stream stream, PdfTextLayoutOptions? options, PdfLogicalMarkdownOptions? markdownOptions, params PdfPageRange[] pageRanges) {
        return ExtractMarkdownByPage(PdfLogicalDocument.LoadPageRanges(stream, options, pageRanges), markdownOptions);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges and concatenates selected pages with Markdown page separators.</summary>
    public static string ExtractMarkdownByPageRangesAsDocument(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractMarkdownByPageRangesAsDocument(stream, (PdfTextLayoutOptions?)null, (PdfLogicalMarkdownOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges and concatenates selected pages with Markdown page separators.</summary>
    public static string ExtractMarkdownByPageRangesAsDocument(Stream stream, PdfTextLayoutOptions? options, PdfLogicalMarkdownOptions? markdownOptions, params PdfPageRange[] pageRanges) {
        return PdfLogicalDocument.LoadPageRanges(stream, options, pageRanges).ToMarkdown(markdownOptions);
    }
    
    /// <summary>Extracts plain text from each page from the current stream position and writes one UTF-8 text file per page.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(Stream stream, string outputDirectory, string baseName = "page", PdfTextLayoutOptions? options = null) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractTextByPage(stream, options);
        return WriteTextPages(baseName, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from the current stream position and writes one UTF-8 text file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(Stream stream, string outputDirectory, string baseName = "page", params PdfPageRange[] pageRanges) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedTextPages(PdfReadDocument.Open(stream), pageRanges);
        return WriteTextPages(baseName, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from the current stream position with layout options and writes one UTF-8 text file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(Stream stream, string outputDirectory, string baseName, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedTextPages(PdfReadDocument.Open(stream), options, pageRanges);
        return WriteTextPages(baseName, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts logical Markdown from each page from the current stream position and writes one UTF-8 Markdown file per page.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPage(Stream stream, string outputDirectory, string baseName = "page", PdfTextLayoutOptions? options = null, PdfLogicalMarkdownOptions? markdownOptions = null) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractMarkdownByPage(stream, options, markdownOptions);
        return WriteMarkdownPages(baseName, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges from the current stream position and writes one UTF-8 Markdown file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPageRanges(Stream stream, string outputDirectory, string baseName = "page", params PdfPageRange[] pageRanges) {
        return ExtractMarkdownByPageRanges(stream, outputDirectory, baseName, null, null, pageRanges);
    }
    
    /// <summary>Extracts logical Markdown from the supplied inclusive one-based page ranges from the current stream position and writes one UTF-8 Markdown file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractMarkdownByPageRanges(Stream stream, string outputDirectory, string baseName, PdfTextLayoutOptions? options, PdfLogicalMarkdownOptions? markdownOptions, params PdfPageRange[] pageRanges) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedMarkdownPages(PdfLogicalDocument.LoadPageRanges(stream, options, pageRanges), markdownOptions);
        return WriteMarkdownPages(baseName, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts plain text from each page from bytes and writes one UTF-8 text file per page.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(byte[] pdf, string outputDirectory, string baseName = "page", PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractTextByPage(pdf, options);
        return WriteTextPages(baseName, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from bytes and writes one UTF-8 text file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(byte[] pdf, string outputDirectory, string baseName = "page", params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedTextPages(PdfReadDocument.Open(pdf), pageRanges);
        return WriteTextPages(baseName, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts plain text from the supplied inclusive one-based page ranges from bytes with layout options and writes one UTF-8 text file per selected source page.</summary>
    public static IReadOnlyList<string> ExtractTextByPageRanges(byte[] pdf, string outputDirectory, string baseName, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
    
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = ExtractSelectedTextPages(PdfReadDocument.Open(pdf), options, pageRanges);
        return WriteTextPages(baseName, fullOutputDirectory, pages);
    }
    
    /// <summary>Extracts structured content for each page from the current stream position.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPage(Stream stream, PdfTextLayoutOptions? options = null) {
        return PdfReadDocument.Open(stream).ExtractStructuredPages(options);
    }
    
    /// <summary>Extracts structured content from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractStructuredByPageRanges(stream, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts structured content from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredPage> ExtractStructuredByPageRanges(Stream stream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        return ExtractStructuredByPageRanges(PdfReadDocument.Open(stream), options, pageRanges);
    }
    
    /// <summary>Extracts detected paragraphs grouped by page from the current stream position.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPage(Stream stream, PdfTextLayoutOptions? options = null) {
        return PdfReadDocument.Open(stream).ExtractParagraphsByPage(options);
    }
    
    /// <summary>Extracts detected paragraphs from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractParagraphsByPageRanges(stream, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected paragraphs from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredParagraphPage> ExtractParagraphsByPageRanges(Stream stream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        return ExtractParagraphsByPageRanges(PdfReadDocument.Open(stream), options, pageRanges);
    }
    
    /// <summary>Extracts detected headings grouped by page from the current stream position.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPage(Stream stream, PdfTextLayoutOptions? options = null) {
        return PdfReadDocument.Open(stream).ExtractHeadingsByPage(options);
    }
    
    /// <summary>Extracts detected headings from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractHeadingsByPageRanges(stream, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected headings from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredHeadingPage> ExtractHeadingsByPageRanges(Stream stream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        return ExtractHeadingsByPageRanges(PdfReadDocument.Open(stream), options, pageRanges);
    }
    
    /// <summary>Extracts detected list items grouped by page from the current stream position.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPage(Stream stream, PdfTextLayoutOptions? options = null) {
        return PdfReadDocument.Open(stream).ExtractListItemsByPage(options);
    }
    
    /// <summary>Extracts detected list items from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractListItemsByPageRanges(stream, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected list items from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredListItemPage> ExtractListItemsByPageRanges(Stream stream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        return ExtractListItemsByPageRanges(PdfReadDocument.Open(stream), options, pageRanges);
    }
    
    /// <summary>Extracts detected tables grouped by page from the current stream position.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPage(Stream stream, PdfTextLayoutOptions? options = null) {
        return PdfReadDocument.Open(stream).ExtractTablesByPage(options);
    }
    
    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractTablesByPageRanges(stream, (PdfTextLayoutOptions?)null, pageRanges);
    }
    
    /// <summary>Extracts detected tables from the supplied inclusive one-based page ranges from the current stream position.</summary>
    public static IReadOnlyList<StructuredTablePage> ExtractTablesByPageRanges(Stream stream, PdfTextLayoutOptions? options, params PdfPageRange[] pageRanges) {
        return ExtractTablesByPageRanges(PdfReadDocument.Open(stream), options, pageRanges);
    }
    
    /// <summary>Extracts plain text from each page from the current stream position using layout options such as column detection and header/footer trimming.</summary>
    public static IReadOnlyList<string> ExtractTextByPage(Stream stream, PdfTextLayoutOptions? options) {
        if (options is null) {
            return ExtractTextByPage(stream);
        }
    
        return ExtractTextByPage(PdfReadDocument.Open(stream), options);
    }
}
