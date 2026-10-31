using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    /// <summary>Checks the emitted artifact rather than trusting removal of the input containers.</summary>
    private static void RequireClassicRedactionOutput(byte[] pdf, PdfLoadOptions? options, CancellationToken cancellationToken) {
        PdfReadDocument output = PdfReadDocument.Open(pdf, options, cancellationToken);
        if (!PdfSyntax.IsClassicRedactionOutput(pdf, output, cancellationToken))
            throw new InvalidOperationException("Redaction did not produce a classic PDF structure.");
    }
}
