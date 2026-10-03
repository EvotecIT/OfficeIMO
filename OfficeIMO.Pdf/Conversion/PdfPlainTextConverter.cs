using System.Text;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Pdf;

/// <summary>Converts literal text through the canonical PDF flow engine without HTML or Markdown parsing.</summary>
public static class PdfPlainTextConverter {
    /// <summary>Reads a caller-owned stream under a byte limit; seekable position is restored and the stream remains open.</summary>
    public static PdfDocumentConversionResult ToPdfDocumentResult(Stream source,
        PdfPlainTextOptions? options = null, long maximumInputBytes = 64L * 1024 * 1024,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(source, nameof(source));
        Guard.Positive(maximumInputBytes, nameof(maximumInputBytes));
        return ToPdfDocumentResult(OfficeStreamReader.ReadAllBytes(source, cancellationToken, maximumInputBytes), options, cancellationToken);
    }
    /// <summary>Decodes bounded bytes using an explicit encoding or a Unicode BOM and strict UTF-8 fallback.</summary>
    public static PdfDocumentConversionResult ToPdfDocumentResult(byte[] bytes,
        PdfPlainTextOptions? options = null, CancellationToken cancellationToken = default) {
        Guard.NotNull(bytes, nameof(bytes));
        PdfPlainTextOptions settings = (options ?? new PdfPlainTextOptions()).Clone();
        cancellationToken.ThrowIfCancellationRequested();
        Encoding encoding = ResolveEncoding(bytes, settings.EncodingName, out int offset);
        int characters = encoding.GetCharCount(bytes, offset, bytes.Length - offset);
        if (characters > settings.MaximumCharacters) throw new InvalidDataException("Plain text exceeds the decoded character limit.");
        return ToPdfDocumentResult(encoding.GetString(bytes, offset, bytes.Length - offset), settings, cancellationToken);
    }

    /// <summary>Preserves source lines, blank lines and spacing; expands tabs at source columns and honors form feeds.</summary>
    public static PdfDocumentConversionResult ToPdfDocumentResult(string text,
        PdfPlainTextOptions? options = null, CancellationToken cancellationToken = default) {
        Guard.NotNull(text, nameof(text));
        PdfPlainTextOptions settings = (options ?? new PdfPlainTextOptions()).Clone();
        if (text.Length > settings.MaximumCharacters) throw new InvalidDataException("Plain text exceeds the decoded character limit.");
        cancellationToken.ThrowIfCancellationRequested();
        var normalized = new StringBuilder(Math.Min(text.Length, settings.MaximumCharacters));
        int column = 0;
        int pageCount = 1;
        for (int index = 0; index < text.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            char value = text[index];
            if (value == '\r') {
                if (index + 1 < text.Length && text[index + 1] == '\n') index++;
                value = '\n';
            }
            if (value == '\f' && ++pageCount > settings.MaximumPages)
                throw new InvalidDataException("Plain text exceeds the explicit page limit.");
            int count = value == '\t' ? settings.TabSize - column % settings.TabSize : 1;
            if (normalized.Length > settings.MaximumCharacters - count)
                throw new InvalidDataException("Plain text exceeds the expanded character limit.");
            if (value == '\t') normalized.Append(' ', count);
            else normalized.Append(value);
            if (value is '\n' or '\f') column = 0;
            else if (!char.IsLowSurrogate(value)) column += count;
        }
        PdfOptions pdfOptions = settings.PdfOptions;
        pdfOptions.PreserveTextWhitespace = true;
        pdfOptions.MaxGeneratedPages = Math.Min(pdfOptions.MaxGeneratedPages ?? settings.MaximumPages, settings.MaximumPages);
        PdfDocument document = PdfDocument.Create(pdfOptions);
        string[] pages = normalized.ToString().Split('\f');
        if (pages.Length > settings.MaximumPages) throw new InvalidDataException("Plain text exceeds the explicit page limit.");
        var style = new PdfParagraphStyle { SpacingBefore = 0, SpacingAfter = 0, LineHeight = 1.2 };
        for (int index = 0; index < pages.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (index > 0) document.Content.PageBreak();
            document.Content.Text(pages[index].Length == 0 ? "\n" : pages[index], style: style);
        }
        return new PdfDocumentConversionResult(document, new PdfConversionReport());
    }

    private static Encoding ResolveEncoding(byte[] bytes, string? name, out int offset) {
        int codePage = 65001;
        offset = 0;
        if (bytes.Length >= 4 && bytes[0] == 0xff && bytes[1] == 0xfe && bytes[2] == 0 && bytes[3] == 0) { codePage = 12000; offset = 4; }
        else if (bytes.Length >= 4 && bytes[0] == 0 && bytes[1] == 0 && bytes[2] == 0xfe && bytes[3] == 0xff) { codePage = 12001; offset = 4; }
        else if (bytes.Length >= 3 && bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf) { offset = 3; }
        else if (bytes.Length >= 2 && bytes[0] == 0xff && bytes[1] == 0xfe) { codePage = 1200; offset = 2; }
        else if (bytes.Length >= 2 && bytes[0] == 0xfe && bytes[1] == 0xff) { codePage = 1201; offset = 2; }
        Encoding encoding = name == null
            ? Encoding.GetEncoding(codePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback)
            : Encoding.GetEncoding(name, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
        if (offset > 0 && encoding.CodePage != codePage)
            throw new InvalidDataException("The selected text encoding conflicts with the byte order mark.");
        return encoding;
    }
}
