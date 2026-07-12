using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>Owns file-source preparation shared by the fluent Markdown-to-PDF extensions.</summary>
internal static class MarkdownPdfConverter {
    internal static PdfCore.PdfDocument ConvertFileMarkdown(string markdown, string fullMarkdownPath, MarkdownPdfSaveOptions options) {
        string? originalBaseDirectory = options.BaseDirectory;
        bool assignedBaseDirectory = string.IsNullOrWhiteSpace(originalBaseDirectory);
        if (assignedBaseDirectory) {
            options.BaseDirectory = Path.GetDirectoryName(fullMarkdownPath);
        }

        try {
            MarkdownDoc document = MarkdownReader.Parse(markdown, MarkdownPdfConverterExtensions.ResolveReaderOptions(options));
            return MarkdownPdfConverterExtensions.ConvertToPdfDocument(document, options);
        } finally {
            if (assignedBaseDirectory) {
                options.BaseDirectory = originalBaseDirectory;
            }
        }
    }
}
