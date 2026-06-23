namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    /// <summary>Loads a PDF from bytes into a typed object model.</summary>
    public static PdfReadDocument Load(byte[] pdf, PdfReadOptions? options = null) {
        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf, options);
        var (map, trailer) = PdfSyntax.ParseObjects(pdf, options);
        return new PdfReadDocument(map, trailer, security, options);
    }

    /// <summary>Loads a PDF from a file path.</summary>
    public static PdfReadDocument Load(string path, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return Load(File.ReadAllBytes(path), options);
    }

    /// <summary>Loads a PDF from the current position of a readable stream.</summary>
    public static PdfReadDocument Load(Stream stream, PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return Load(buffer.ToArray(), options);
    }

    /// <summary>Extracts full‑document plain text (pages separated by blank lines).</summary>
    public string ExtractText() {
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < Pages.Count; i++) {
            if (i > 0) sb.AppendLine();
            sb.Append(Pages[i].ExtractText());
        }
        return sb.ToString();
    }

    /// <summary>Extracts image XObjects from all pages in page order.</summary>
    public IReadOnlyList<PdfExtractedImage> ExtractImages() => PdfImageExtractor.ExtractImages(this);

    /// <summary>Extracts embedded file attachments from the document catalog.</summary>
    public IReadOnlyList<PdfExtractedAttachment> ExtractAttachments() => PdfAttachmentExtractor.ExtractAttachments(_objects, _trailerRaw);
}
