using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Routes printer delivery through the capabilities of the installed Studio edition.</summary>
internal static class StudioPrinterService {
    internal static IPdfPrinterService Create() => StudioDistributionPolicy.ExternalToolsAllowed
        ? new PdfPrinterService() : new PreviewOnlyPrinterService();

    private sealed class PreviewOnlyPrinterService : IPdfPrinterService {
        private static NotSupportedException Unavailable() => new(
            "Printer queue delivery requires an external system tool and is unavailable in the Mac App Store edition. Save the prepared PDF and print it from a system application.");
        public Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CancellationToken cancellationToken = default) => Task.FromException<IReadOnlyList<PdfPrinterInfo>>(Unavailable());
        public Task<IReadOnlyList<PdfPaperSourceInfo>> GetPaperSourcesAsync(string printerName, CancellationToken cancellationToken = default) => Task.FromException<IReadOnlyList<PdfPaperSourceInfo>>(Unavailable());
        public Task<PdfPrintSubmission> SubmitAsync(PdfPreparedPrintDocument document, PdfPrintDeliveryOptions options, CancellationToken cancellationToken = default) => Task.FromException<PdfPrintSubmission>(Unavailable());
    }
}
