using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    private CompletedOutput? _completedOutput;

    private PdfCreateOptions OptionsToChange {
        get {
            Volatile.Write(ref _completedOutput, null);
            return _options;
        }
    }

    /// <summary>
    /// Keeps bytes a conversion has already written and verified, so saving the unchanged model
    /// does not repeat that work. Any change to the model discards them; each is published once.
    /// </summary>
    internal void RetainCompletedOutput(byte[] bytes, PdfConversionReport diagnostics) =>
        Volatile.Write(ref _completedOutput, new CompletedOutput(bytes, diagnostics));

    private bool TryPublishCompletedOutput(PdfSaveOptions? options, out byte[] bytes) {
        bytes = Array.Empty<byte>();
        if (options?.Encryption != null || Interlocked.Exchange(ref _completedOutput, null) is not { } completed) return false;
        if (options?.AllowLoss != true) PdfConversionEvidence.CreateReport("Pdf", new[] { completed.Diagnostics }).ThrowIfLoss();
        bytes = completed.Bytes;
        return true;
    }

    private sealed class CompletedOutput {
        internal CompletedOutput(byte[] bytes, PdfConversionReport diagnostics) {
            Bytes = bytes;
            Diagnostics = diagnostics;
        }

        internal byte[] Bytes { get; }
        internal PdfConversionReport Diagnostics { get; }
    }
}
