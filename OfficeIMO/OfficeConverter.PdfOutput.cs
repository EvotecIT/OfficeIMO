using System.Diagnostics;
using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private static async Task<OfficeConversionReport> ConvertCapturedAsync(Stream input, OfficeBoundedMemoryStream output,
        OfficeConversionOptions settings, CancellationToken token) {
        if (!settings.CompressPdfOutput) return await ConvertCapturedCoreAsync(input, output, settings, token).ConfigureAwait(false);
        if (FormatForExtension(Path.GetExtension(settings.TargetName)) != OfficeDocumentFormat.Pdf)
            throw new ArgumentException("PDF compression requires a PDF destination.", nameof(settings));
        var timer = Stopwatch.StartNew();
        PdfCreateOptions? creation = settings.Word?.PdfOptions ?? settings.Excel?.PdfOptions ?? settings.PowerPoint?.PdfOptions
            ?? settings.Markdown?.PdfOptions ?? settings.Html?.PdfOptions ?? settings.PlainText?.PdfOptions ?? settings.ReaderPdf?.PdfOptions;
        if (Registrations.TryGetValue(SelectRoute(settings).Id, out IOfficeConversionRoute? route) && route is IOfficeConversionOutputValidation owner)
            creation = owner.GetOutputCreationOptions(settings) ?? creation;
        PdfStandardEncryptionOptions? encryption = creation?.Encryption?.Clone();
        if (encryption != null) creation!.ClearEncryption();
        OfficeConversionReport report = await ConvertCapturedCoreAsync(input, output, settings, token).ConfigureAwait(false);
        try {
        PdfOptimizationActionResult optimized = CompressPdfBytes(output.ToArray(), settings.MaximumOutputBytes, null, token);
        report = new OfficeConversionReport(report.Source, report.Target, itemsWritten: report.ItemsWritten, elapsed: timer.Elapsed,
            diagnostics: report.Diagnostics.Concat(new[] { new OfficeDiagnostic(OfficeDiagnosticCodes.PdfOutputCompressed,
                OfficeDiagnosticSeverity.Information, "Verified lossless PDF compression completed; saved " + optimized.SavedBytes + " bytes.", format: "Pdf") }));
        byte[] bytes = optimized.Bytes;
        if (encryption != null) {
            PdfSecurityMutationResult encrypted = PdfSecurityEditor.Encrypt(bytes, encryption,
                maximumOutputBytes: settings.MaximumOutputBytes, cancellationToken: token);
            if (!encrypted.PreservationReport.IsPreserved) throw new InvalidOperationException("PDF encryption did not preserve the converted document.");
            bytes = encrypted.Pdf;
        }
        token.ThrowIfCancellationRequested();
        output.Position = 0;
        output.SetLength(0);
        await output.WriteAsync(bytes, 0, bytes.Length, token).ConfigureAwait(false);
        return new OfficeConversionReport(report.Source, report.Target, itemsWritten: report.ItemsWritten, elapsed: timer.Elapsed,
            diagnostics: report.Diagnostics);
        } catch (Exception error) when (error is not OperationCanceledException and not OfficeConversionException and not OutOfMemoryException and not StackOverflowException) {
            throw new OfficeConversionException(error.Message, report, error);
        }
    }

    internal static PdfOptimizationActionResult CompressPdfBytes(byte[] bytes, long maximumOutputBytes,
        PdfLoadOptions? readOptions, CancellationToken token) {
        var compression = PdfOptimizationOptions.Create(PdfOptimizationProfile.MaximumCompression);
        compression.KeepOriginalWhenNotSmaller = true;
        compression.CancellationToken = token;
        compression.MaximumOutputBytes = maximumOutputBytes;
        using var document = PdfDocument.Load(new MemoryStream(bytes), readOptions, token);
        PdfOptimizationActionResult optimized = document.Optimization.Apply(compression);
        if (!optimized.PreservationReport.IsPreserved)
            throw new InvalidOperationException("PDF compression did not preserve the converted document.");
        return optimized;
    }
}
