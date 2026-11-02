using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    private void WriteNativeWithLossCheck(Stream destination, PdfSaveOptions? options, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        ThrowIfTextEncodingPreflightFails(token);
        if (options?.Encryption != null) {
            byte[] bytes = ToBytes(options, token);
            token.ThrowIfCancellationRequested();
            destination.Write(bytes, 0, bytes.Length);
            token.ThrowIfCancellationRequested();
            return;
        }
        using var scope = new PdfOutputDiagnosticScope();
        WritePdfCore(destination, token);
        if (options?.AllowLoss != true) PdfConversionEvidence.CreateReport("Pdf", new[] { scope.Report }).ThrowIfLoss();
    }

    private void SaveNativeStream(Stream destination, PdfSaveOptions? options, CancellationToken token) {
        if (options?.Encryption == null && (_source != null || options?.AllowLoss == true)) {
            _ = SaveWithEvidence(destination, token);
            return;
        }
        token.ThrowIfCancellationRequested();
        string directory = OfficeTemporaryDirectory.Create("officeimo-pdf-output-");
        try {
            using var staged = new FileStream(System.IO.Path.Combine(directory, "document.pdf"), FileMode.CreateNew,
                FileAccess.ReadWrite, FileShare.None, 81920, FileOptions.SequentialScan);
            WriteNativeWithLossCheck(staged, options, token);
            staged.Position = 0;
            OfficeStreamWriter.Write(destination, output => CopyValidatedOutput(staged, output, token));
        } finally { Directory.Delete(directory, recursive: true); }
    }

    private async Task SaveNativeStreamAsync(Stream destination, PdfSaveOptions? options, CancellationToken token) {
        if (options?.Encryption == null && (_source != null || options?.AllowLoss == true)) {
            _ = await SaveWithEvidenceAsync(destination, token).ConfigureAwait(false);
            return;
        }
        token.ThrowIfCancellationRequested();
        string directory = OfficeTemporaryDirectory.Create("officeimo-pdf-output-");
        try {
            using var staged = new FileStream(System.IO.Path.Combine(directory, "document.pdf"), FileMode.CreateNew,
                FileAccess.ReadWrite, FileShare.None, 81920, FileOptions.Asynchronous | FileOptions.SequentialScan);
            WriteNativeWithLossCheck(staged, options, token);
            staged.Position = 0;
            if (PdfComplianceValidator.RequiresExactArtifactValidation(_options)) {
                await OfficeStreamWriter.WriteAsync(destination,
                    output => CopyValidatedOutput(staged, output, CancellationToken.None), token,
                    commitIsNonCancellable: true).ConfigureAwait(false);
            } else {
                await OfficeStreamWriter.WriteAsync(destination,
                    (output, cancellation) => staged.CopyToAsync(output, 81920, cancellation), token).ConfigureAwait(false);
            }
        } finally { Directory.Delete(directory, recursive: true); }
    }

    private static void CopyValidatedOutput(Stream source, Stream destination, CancellationToken token) {
        byte[] buffer = new byte[81920];
        int count;
        while ((count = source.Read(buffer, 0, buffer.Length)) != 0) {
            token.ThrowIfCancellationRequested();
            destination.Write(buffer, 0, count);
            token.ThrowIfCancellationRequested();
        }
    }
}
