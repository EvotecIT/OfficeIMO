using OfficeIMO.Drawing.Internal;
namespace OfficeIMO.Pdf;

internal static partial class PdfIncrementalUpdater {
    /// <summary>Prepares byte ranges, calls caller-owned key infrastructure, and applies its signature container.</summary>
    public static PdfExternalSignatureCompletion SignExternal(
        byte[] pdf,
        IPdfExternalSigner signer,
        PdfExternalSignatureOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(signer, nameof(signer));
        if (string.IsNullOrWhiteSpace(signer.Name)) {
            throw new ArgumentException("External signer name cannot be empty.", nameof(signer));
        }

        PdfExternalSignaturePreparation preparation = PrepareExternalSignature(pdf, options);
        byte[] signatureContents = signer.Sign(new PdfExternalSignatureRequest(preparation));
        if (signatureContents is null || signatureContents.Length == 0) {
            throw new InvalidOperationException(signer.Name + " returned empty signature contents.");
        }

        byte[] completedPdf = ApplyExternalSignature(preparation, signatureContents);
        return new PdfExternalSignatureCompletion(
            completedPdf,
            preparation,
            signer.Name,
            signatureContents.Length);
    }

    /// <summary>Signs a PDF from a readable stream through caller-owned key infrastructure.</summary>
    public static PdfExternalSignatureCompletion SignExternal(
        Stream input,
        IPdfExternalSigner signer,
        PdfExternalSignatureOptions? options = null) {
        Guard.NotNull(input, nameof(input));
        if (!input.CanRead) throw new ArgumentException("Stream must be readable.", nameof(input));
        using var buffer = new MemoryStream();
        input.CopyTo(buffer);
        return SignExternal(buffer.ToArray(), signer, options);
    }

    /// <summary>Signs a PDF file through caller-owned key infrastructure and writes the completed output.</summary>
    public static PdfExternalSignatureCompletion SignExternal(
        string inputPath,
        string outputPath,
        IPdfExternalSigner signer,
        PdfExternalSignatureOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNullOrWhiteSpace(outputPath, nameof(outputPath));
        PdfExternalSignatureCompletion completion = SignExternal(File.ReadAllBytes(inputPath), signer, options);
        OfficeFileCommit.WriteAllBytes(outputPath, completion.Pdf);
        return completion;
    }
}
