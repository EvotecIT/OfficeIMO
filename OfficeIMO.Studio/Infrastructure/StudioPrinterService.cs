using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Creates the printer delivery service for the desktop editions.</summary>
internal static class StudioPrinterService {
    internal static IPdfPrinterService Create() => new PdfPrinterService();
}
