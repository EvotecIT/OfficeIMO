using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// Explicit Markdown-to-PDF conversion entry points for scenarios where the source is a file path.
/// </summary>
public static class MarkdownPdfConverter {
    /// <summary>
    /// Converts a Markdown file to a first-party OfficeIMO PDF document model.
    /// </summary>
    public static PdfCore.PdfDoc FromFile(string markdownPath, MarkdownPdfSaveOptions? options = null) {
        if (string.IsNullOrWhiteSpace(markdownPath)) {
            throw new ArgumentException("Markdown file path cannot be empty.", nameof(markdownPath));
        }

        options ??= new MarkdownPdfSaveOptions();
        string fullPath = Path.GetFullPath(markdownPath);
        string markdown = File.ReadAllText(fullPath, Encoding.UTF8);
        return ConvertFileMarkdown(markdown, fullPath, options);
    }

    /// <summary>
    /// Converts a Markdown file to PDF bytes.
    /// </summary>
    public static byte[] SaveFileAsPdf(string markdownPath, MarkdownPdfSaveOptions? options = null) {
        return FromFile(markdownPath, options).ToBytes();
    }

    /// <summary>
    /// Saves a Markdown file as a PDF file.
    /// </summary>
    public static void SaveFileAsPdf(string markdownPath, string pdfPath, MarkdownPdfSaveOptions? options = null) {
        FromFile(markdownPath, options).Save(pdfPath);
    }

    /// <summary>
    /// Writes a Markdown file as PDF to a stream.
    /// </summary>
    public static void SaveFileAsPdf(string markdownPath, Stream stream, MarkdownPdfSaveOptions? options = null) {
        FromFile(markdownPath, options).Save(stream);
    }

    internal static PdfCore.PdfDoc ConvertFileMarkdown(string markdown, string fullMarkdownPath, MarkdownPdfSaveOptions options) {
        string? originalBaseDirectory = options.BaseDirectory;
        bool assignedBaseDirectory = string.IsNullOrWhiteSpace(originalBaseDirectory);
        if (assignedBaseDirectory) {
            options.BaseDirectory = Path.GetDirectoryName(fullMarkdownPath);
        }

        try {
            return markdown.ToPdfDocument(options);
        } finally {
            if (assignedBaseDirectory) {
                options.BaseDirectory = originalBaseDirectory;
            }
        }
    }
}
