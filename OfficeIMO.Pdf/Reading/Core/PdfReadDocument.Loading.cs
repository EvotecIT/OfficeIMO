namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    /// <summary>Loads a PDF from bytes into a typed object model.</summary>
    public static PdfReadDocument Load(byte[] pdf, PdfReadOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf, options);
        var (map, trailer) = PdfSyntax.ParseObjects(pdf, options, out PdfRepairReport repairReport);
        if (options?.Password is not null && security.HasEncryption) {
            security = PdfSyntax.ReadDocumentSecurityInfo(pdf, map, trailer, security);
        }

        return new PdfReadDocument(map, trailer, security, repairReport, options);
    }

    /// <summary>Loads a PDF from a file path.</summary>
    public static PdfReadDocument Load(string path, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        PdfReadLimits limits = GetValidatedLimits(options);
        var file = new FileInfo(path);
        if (file.Length > limits.MaxInputBytes) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.InputBytes, limits.MaxInputBytes, file.Length);
        }

        return Load(File.ReadAllBytes(path), options);
    }

    /// <summary>Loads a PDF from the current position of a readable stream.</summary>
    public static PdfReadDocument Load(Stream stream, PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        PdfReadLimits limits = GetValidatedLimits(options);
        if (stream.CanSeek) {
            long remaining = stream.Length - stream.Position;
            if (remaining > limits.MaxInputBytes) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.InputBytes, limits.MaxInputBytes, remaining);
            }
        }

        using var buffer = new MemoryStream();
        var chunk = new byte[81920];
        int read;
        while ((read = stream.Read(chunk, 0, chunk.Length)) > 0) {
            long nextLength = buffer.Length + read;
            if (nextLength > limits.MaxInputBytes) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.InputBytes, limits.MaxInputBytes, nextLength);
            }

            buffer.Write(chunk, 0, read);
        }

        return Load(buffer.ToArray(), options);
    }

    private static PdfReadLimits GetValidatedLimits(PdfReadOptions? options) {
        PdfReadLimits limits = options?.Limits ?? new PdfReadLimits();
        limits.Validate();
        return limits;
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
    public IReadOnlyList<PdfExtractedAttachment> ExtractAttachments() => PdfAttachmentExtractor.ExtractAttachments(_objects, _trailerRaw, _options.Limits);
}
