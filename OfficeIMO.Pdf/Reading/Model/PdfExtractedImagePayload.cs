namespace OfficeIMO.Pdf;

/// <summary>The extracted bytes of an image and the facts that depend on how they were produced.</summary>
internal sealed class PdfExtractedImagePayload {
    internal PdfExtractedImagePayload(
        byte[] bytes,
        bool ownsBytes,
        string? fileExtension,
        string? mimeType,
        bool isImageFile,
        string? transparencyMaskKind,
        bool transparencyMaskResolved) {
        // A payload built for this image, such as a normalized PNG, is not shared; source stream bytes are.
        Bytes = ownsBytes ? bytes : (byte[])bytes.Clone();
        Facts = new PdfExtractedImagePayloadFacts(fileExtension, mimeType, isImageFile, transparencyMaskKind, transparencyMaskResolved);
    }

    internal PdfExtractedImagePayload(byte[] ownedBytes, PdfExtractedImagePayloadFacts facts) {
        Bytes = ownedBytes;
        Facts = facts;
    }

    internal byte[] Bytes { get; }
    internal PdfExtractedImagePayloadFacts Facts { get; }
}

/// <summary>What an extracted payload is, kept apart from its bytes so a deferred image can answer without holding them.</summary>
internal sealed class PdfExtractedImagePayloadFacts {
    internal PdfExtractedImagePayloadFacts(
        string? fileExtension,
        string? mimeType,
        bool isImageFile,
        string? transparencyMaskKind,
        bool transparencyMaskResolved) {
        FileExtension = fileExtension;
        MimeType = mimeType;
        IsImageFile = isImageFile;
        TransparencyMaskKind = transparencyMaskKind;
        TransparencyMaskResolved = transparencyMaskResolved;
    }

    internal string? FileExtension { get; }
    internal string? MimeType { get; }
    internal bool IsImageFile { get; }
    internal string? TransparencyMaskKind { get; }
    internal bool TransparencyMaskResolved { get; }
}
