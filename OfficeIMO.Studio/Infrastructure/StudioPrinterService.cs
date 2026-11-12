using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Routes printer delivery through the capabilities of the installed Studio edition.</summary>
internal static class StudioPrinterService {
    internal static IPdfPrinterService Create() => StudioDistributionPolicy.ExternalToolsAllowed
        ? new PdfPrinterService() : new PreviewOnlyPrinterService();

    private sealed class PreviewOnlyPrinterService : IPdfPrinterService {
        private static NotSupportedException Unavailable() => new(
            OperatingSystem.IsIOS() ? "Save the prepared PDF and use Share to print with iOS. Desktop printer queue delivery cannot run on iPhone or iPad." :
            "Printer queue delivery requires an external system tool and is unavailable in the Mac App Store edition. Save the prepared PDF and print it from a system application.");
        public Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CancellationToken cancellationToken = default) => Task.FromException<IReadOnlyList<PdfPrinterInfo>>(Unavailable());
        public Task<PdfPrinterCapabilities> GetCapabilitiesAsync(string printerName, CancellationToken cancellationToken = default) => Task.FromException<PdfPrinterCapabilities>(Unavailable());
        public Task<PdfPrintSubmission> SubmitAsync(PdfPrintJob job, PdfPrintDeliveryOptions options, CancellationToken cancellationToken = default) => Task.FromException<PdfPrintSubmission>(Unavailable());
    }
}
