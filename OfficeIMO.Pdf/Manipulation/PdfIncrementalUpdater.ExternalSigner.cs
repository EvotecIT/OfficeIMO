using OfficeIMO.Drawing.Internal;
namespace OfficeIMO.Pdf;

internal static partial class PdfIncrementalUpdater {
    /// <summary>Prepares byte ranges, calls caller-owned key infrastructure, and applies its signature container.</summary>
    public static PdfExternalSignatureCompletion SignExternal(
        byte[] pdf,
        IPdfExternalSigner signer,
        PdfExternalSignatureOptions? options = null) =>
        SignExternal(pdf, signer, options, readOptions: null);

    internal static PdfExternalSignatureCompletion SignExternal(
        byte[] pdf,
        IPdfExternalSigner signer,
        PdfExternalSignatureOptions? options,
        PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(signer, nameof(signer));
        PdfExternalSignatureOptions effectiveOptions = options ?? new PdfExternalSignatureOptions();
        ValidateSigningInput(pdf.LongLength, effectiveOptions);
        effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(signer.Name)) {
            throw new ArgumentException("External signer name cannot be empty.", nameof(signer));
        }

        PdfExternalSignaturePreparation preparation = PrepareExternalSignature(pdf, effectiveOptions, readOptions);
        effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
        byte[] signatureContents = signer.Sign(new PdfExternalSignatureRequest(preparation));
        if (signatureContents is null || signatureContents.Length == 0) {
            throw new InvalidOperationException(signer.Name + " returned empty signature contents.");
        }

        byte[] completedPdf = ApplyExternalSignature(preparation, signatureContents);
        return new PdfExternalSignatureCompletion(
            completedPdf,
            preparation,
            signer.Name,
            signatureContents.Length,
            preparation.GetCompletionReadOptions(completedPdf.LongLength));
    }

    /// <summary>Signs a PDF from a readable stream through caller-owned key infrastructure.</summary>
    public static PdfExternalSignatureCompletion SignExternal(
        Stream input,
        IPdfExternalSigner signer,
        PdfExternalSignatureOptions? options = null) {
        Guard.NotNull(input, nameof(input));
        if (!input.CanRead) throw new ArgumentException("Stream must be readable.", nameof(input));
        PdfExternalSignatureOptions effectiveOptions = options ?? new PdfExternalSignatureOptions();
        byte[] pdf = ReadSigningInput(input, effectiveOptions);
        return SignExternal(pdf, signer, effectiveOptions);
    }

    private static byte[] ReadSigningInput(Stream input, PdfExternalSignatureOptions effectiveOptions) {
        ValidateSigningInput(0, effectiveOptions);
        try {
            return OfficeStreamReader.ReadRemainingBytes(
                input,
                effectiveOptions.CancellationToken,
                effectiveOptions.MaxInputBytes);
        } catch (InvalidDataException) {
            long observedBytes = input.CanSeek
                ? Math.Max(0L, input.Length - input.Position)
                : checked(effectiveOptions.MaxInputBytes + 1L);
            throw PdfReadLimitException.Create(PdfReadLimitKind.InputBytes,
                effectiveOptions.MaxInputBytes, observedBytes);
        }
    }

    /// <summary>Signs a PDF file through caller-owned key infrastructure and writes the completed output.</summary>
    public static PdfExternalSignatureCompletion SignExternal(
        string inputPath,
        string outputPath,
        IPdfExternalSigner signer,
        PdfExternalSignatureOptions? options = null) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNullOrWhiteSpace(outputPath, nameof(outputPath));
        PdfExternalSignatureOptions effectiveOptions = options ?? new PdfExternalSignatureOptions();
        PdfExternalSignatureCompletion completion;
        using (var input = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read)) {
            completion = SignExternal(input, signer, effectiveOptions);
        }
        OfficeFileCommit.WriteAllBytes(outputPath, completion.Pdf);
        return completion;
    }

    private static void ValidateSigningInput(long observedBytes, PdfExternalSignatureOptions options) {
        if (options.MaxInputBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), options.MaxInputBytes,
                "Maximum signing input bytes must be positive.");
        }
        if (observedBytes > options.MaxInputBytes) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.InputBytes, options.MaxInputBytes, observedBytes);
        }
    }
}
