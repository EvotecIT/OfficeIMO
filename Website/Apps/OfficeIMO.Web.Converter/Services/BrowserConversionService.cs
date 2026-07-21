using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Web.Converter.Services;

public sealed class BrowserConversionService {
    public const long MaxPackageBytes = 25L * 1024L * 1024L;
    public const int MaxTextInputChars = 500_000;

    public ConversionResult ConvertFile(ConversionRoute route, SelectedDocument file, bool fastPreview) => route.Id switch {
        "docx-pdf" => ConvertWordToPdf(file),
        "xlsx-pdf" => ConvertExcelToPdf(file, fastPreview),
        "pptx-pdf" => ConvertPowerPointToPdf(file),
        _ => throw new NotSupportedException($"The route '{route.Id}' does not accept a document upload.")
    };

    public ConversionResult ConvertText(ConversionRoute route, string input) {
        ArgumentNullException.ThrowIfNull(input);
        if (input.Length > MaxTextInputChars) {
            throw new ArgumentOutOfRangeException(nameof(input),
                $"Text input is limited to {MaxTextInputChars:N0} characters in the browser converter.");
        }

        return route.Id switch {
            "markdown-html" => ConvertMarkdownToHtml(input),
            "html-markdown" => ConvertHtmlToMarkdown(input),
            "markdown-docx" => ConvertMarkdownToWord(input),
            _ => throw new NotSupportedException($"The route '{route.Id}' does not accept text input.")
        };
    }

    private static ConversionResult ConvertWordToPdf(SelectedDocument file) {
        using var stream = new MemoryStream(file.Bytes, writable: false);
        using WordDocument document = WordDocument.Load(stream,
            new WordLoadOptions {
                AccessMode = DocumentAccessMode.ReadOnly,
                PackageSecurity = CreateBrowserPackageSecurity()
            });
        var options = new PdfSaveOptions {
            Title = Path.GetFileNameWithoutExtension(file.Name),
            IncludePageNumbers = true
        };
        var conversion = document.ToPdfDocumentResult(options);
        byte[] bytes = conversion.Value.ToBytes();
        return PdfResult(file, bytes, conversion.Warnings.Select(static warning => warning.ToString()).ToArray());
    }

    private static ConversionResult ConvertExcelToPdf(SelectedDocument file, bool fastPreview) {
        using var stream = new MemoryStream(file.Bytes, writable: false);
        using ExcelDocument document = ExcelDocument.Load(stream,
            new ExcelLoadOptions {
                AccessMode = DocumentAccessMode.ReadOnly,
                PackageSecurity = CreateBrowserPackageSecurity()
            });
        var options = new ExcelPdfSaveOptions { MaxRowsPerSheet = fastPreview ? 250 : null };
        var conversion = document.ToPdfDocumentResult(options);
        byte[] bytes = conversion.Value.ToBytes();
        return PdfResult(file, bytes, conversion.Warnings.Select(static warning => warning.ToString()).ToArray());
    }

    private static ConversionResult ConvertPowerPointToPdf(SelectedDocument file) {
        using var stream = new MemoryStream(file.Bytes, writable: false);
        using PowerPointPresentation presentation = PowerPointPresentation.Load(stream,
            new PowerPointLoadOptions {
                AccessMode = DocumentAccessMode.ReadOnly,
                PackageSecurity = CreateBrowserPackageSecurity()
            });
        var options = new PowerPointPdfSaveOptions { WarnOnPictureAspectRatioDistortion = true };
        var conversion = presentation.ToPdfDocumentResult(options);
        byte[] bytes = conversion.Value.ToBytes();
        return PdfResult(file, bytes, conversion.Warnings.Select(static warning => warning.ToString()).ToArray());
    }

    private static ConversionResult PdfResult(SelectedDocument file, byte[] bytes, IReadOnlyList<string> warnings) {
        EnsurePdf(bytes);
        return new ConversionResult(
            bytes,
            Path.GetFileNameWithoutExtension(file.Name) + ".pdf",
            "application/pdf",
            null,
            null,
            warnings);
    }

    private static ConversionResult ConvertMarkdownToHtml(string input) {
        var options = MarkdownRendererPresets.CreateStrictMinimalPortable();
        options.MarkdownOverflowHandling = OverflowHandling.Throw;
        string html = global::OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(input, options);
        return TextResult(html, "officeimo-markdown.html", "text/html;charset=utf-8", html);
    }

    private static ConversionResult ConvertHtmlToMarkdown(string input) {
        var options = HtmlToMarkdownOptions.CreatePortableProfile();
        string markdown = HtmlConversionDocument.Parse(input).ToMarkdown(options);
        return TextResult(markdown, "officeimo-html.md", "text/markdown;charset=utf-8", null);
    }

    private static ConversionResult ConvertMarkdownToWord(string input) {
        MarkdownDoc markdown = MarkdownReader.Parse(input);
        using WordDocument document = markdown.ToWordDocument(new MarkdownToWordOptions {
            PreferNarrativeSingleLineDefinitions = true
        });
        byte[] bytes = document.ToBytes();
        return new ConversionResult(
            bytes,
            "officeimo-markdown.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            $"DOCX generated: {FormatBytes(bytes.Length)}.",
            null,
            []);
    }

    private static ConversionResult TextResult(string text, string fileName, string contentType, string? htmlPreview) =>
        new(Encoding.UTF8.GetBytes(text), fileName, contentType, text, htmlPreview, []);

    private static OfficePackageSecurityOptions CreateBrowserPackageSecurity() {
        OfficePackageSecurityOptions options = OfficePackageSecurityOptions.SecureDefaults;
        options.MaxPackageBytes = MaxPackageBytes;
        options.MaxPartCount = 5_000;
        options.MaxPartUncompressedBytes = 32L * 1024L * 1024L;
        options.MaxTotalUncompressedBytes = 128L * 1024L * 1024L;
        options.MaxCompressionRatio = 200D;
        return options;
    }

    private static void EnsurePdf(byte[] bytes) {
        if (bytes.Length < 4 || bytes[0] != 0x25 || bytes[1] != 0x50 || bytes[2] != 0x44 || bytes[3] != 0x46) {
            throw new InvalidDataException("The conversion did not return a valid PDF header.");
        }
    }

    private static string FormatBytes(long bytes) {
        string[] units = ["B", "KB", "MB", "GB"];
        double value = bytes;
        int unit = 0;
        while (value >= 1024 && unit < units.Length - 1) {
            value /= 1024;
            unit++;
        }
        return unit == 0 ? $"{bytes} B" : $"{value:0.##} {units[unit]}";
    }
}
