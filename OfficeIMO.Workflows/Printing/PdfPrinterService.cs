using OfficeIMO.Internal;
namespace OfficeIMO.Workflows;

/// <summary>An installed operating-system printer queue.</summary>
public sealed record PdfPrinterInfo(string Name, bool IsDefault, bool RequiresOutputFile) {
    /// <inheritdoc />
    public override string ToString() => Name;
}

/// <summary>A paper source reported by a specific printer queue.</summary>
/// <param name="Id">Provider identifier, valid only for the queried queue.</param>
/// <param name="Name">Driver-supplied display name.</param>
public sealed record PdfPaperSourceInfo(string Id, string Name) {
    /// <inheritdoc />
    public override string ToString() => Name;
}

/// <summary>Driver duplex request.</summary>
public enum PdfPrintDuplex {
    /// <summary>Keep the printer's configured duplex default.</summary>
    PrinterDefault,
    /// <summary>Print one side per sheet.</summary>
    SingleSided,
    /// <summary>Turn double-sided pages along the long edge.</summary>
    LongEdge,
    /// <summary>Turn double-sided pages along the short edge.</summary>
    ShortEdge
}

/// <summary>Delivery options applied to a composed print job.</summary>
public sealed class PdfPrintDeliveryOptions {
    /// <summary>Installed queue name, or an <c>ipp://</c> or <c>ipps://</c> printer URI.</summary>
    public required string PrinterName { get; set; }
    /// <summary>Spooler document label.</summary>
    public string DocumentName { get; set; } = "OfficeIMO document";
    /// <summary>Number of collated copies, from 1 through 100.</summary>
    public int Copies { get; set; } = 1;
    /// <summary>Duplex request, or the driver default.</summary>
    public PdfPrintDuplex Duplex { get; set; }
    /// <summary>Identifier returned by paper-source discovery, or null to retain the queue default.</summary>
    public string? PaperSourceId { get; set; }
    /// <summary>New local output path for a Windows file printer. Existing paths are rejected before submission.</summary>
    public string? OutputFilePath { get; set; }
    /// <summary>Highest resolution, from 72 through 600 DPI, used when a transport rasterizes sheets. Defaults to 300.</summary>
    /// <remarks>Windows rasterizes at the device resolution up to this limit. IPP printers and CUPS queues receive raster sheets only for grayscale or raster jobs, or when the printer does not accept PDF.</remarks>
    public double MaximumRasterDpi { get; set; } = 300;

    internal PdfPrintDeliveryOptions Snapshot() {
        if ((PrinterName == null || string.IsNullOrWhiteSpace(PrinterName))) throw new ArgumentException("A non-empty value is required.", nameof(PrinterName));
        if ((DocumentName == null || string.IsNullOrWhiteSpace(DocumentName))) throw new ArgumentException("A non-empty value is required.", nameof(DocumentName));
        if (PrinterName.IndexOf('\0') >= 0 || DocumentName.IndexOf('\0') >= 0) throw new ArgumentException("Printer and document names cannot contain null characters.");
        _ = PdfPrinterAddress.TryGetIppUri(PrinterName, nameof(PrinterName), out _);
        if (Copies is < 1 or > 100) throw new ArgumentOutOfRangeException(nameof(Copies));
        if (!OfficeRuntimeCompatibility.IsDefined(Duplex)) throw new ArgumentOutOfRangeException(nameof(Duplex));
        if (!OfficeRuntimeCompatibility.IsFinite(MaximumRasterDpi) || MaximumRasterDpi is < 72 or > 600) throw new ArgumentOutOfRangeException(nameof(MaximumRasterDpi));
        if (PaperSourceId is not null && ((PaperSourceId == null || string.IsNullOrWhiteSpace(PaperSourceId)) || PaperSourceId.Length > 256 || PaperSourceId.Any(char.IsControl)))
            throw new ArgumentException("The paper-source identifier is invalid.", nameof(PaperSourceId));
        return new() { PrinterName = PrinterName, DocumentName = DocumentName, Copies = Copies, Duplex = Duplex, PaperSourceId = PaperSourceId, OutputFilePath = OutputFilePath, MaximumRasterDpi = MaximumRasterDpi };
    }
}

/// <summary>Receipt proving that a job was accepted by a print queue or IPP printer.</summary>
/// <remarks>Acceptance does not prove that paper or a virtual-printer file has finished printing.</remarks>
public sealed record PdfPrintSubmission(string PrinterName, string JobId, int SheetCount, int Copies, string? OutputFilePath) {
    /// <summary>Local staging cleanup problem after acceptance; the receipt remains valid and the job must not be resubmitted.</summary>
    public string? CleanupWarning { get; init; }
    /// <summary>Requested job attributes, such as <c>sides</c> or <c>media</c>, that the printer ignored, substituted or does not support.</summary>
    public IReadOnlyList<string> IgnoredAttributes { get; init; } = [];
}

/// <summary>A delivery failure after submission began; retrying may print duplicate pages.</summary>
public sealed class PdfPrintDeliveryException : IOException {
    /// <summary>Creates an uncertain-delivery result for a printer provider after submission started.</summary>
    public PdfPrintDeliveryException(string? jobId, Exception inner) : base(
        "Printer delivery was interrupted after submission began. Check the printer queue before retrying; some pages may already have printed.", inner) => JobId = jobId;
    /// <summary>Known spooler job identifier, or null if the acknowledgement was interrupted.</summary>
    public string? JobId { get; }
}

/// <summary>Printer discovery and job delivery. Transports receive the job content and the delivery options.</summary>
/// <remarks>Every member that takes a printer name also accepts an <c>ipp://</c> or <c>ipps://</c> printer URI.</remarks>
public interface IPdfPrinterService {
    /// <summary>Lists installed queues without changing printer settings.</summary>
    Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CancellationToken cancellationToken = default);
    /// <summary>Reports the selected printer's media, trays and features without changing printer settings.</summary>
    Task<PdfPrinterCapabilities> GetCapabilitiesAsync(string printerName, CancellationToken cancellationToken = default);
    /// <summary>Submits a composed job and returns an acceptance receipt.</summary>
    Task<PdfPrintSubmission> SubmitAsync(PdfPrintJob job, PdfPrintDeliveryOptions options, CancellationToken cancellationToken = default);
}

/// <summary>Delivers print jobs to IPP printers by URI, to CUPS queues through the local scheduler, and to Windows queues through GDI.</summary>
/// <remarks>
/// IPP printers receive vector PDF when they accept <c>application/pdf</c>, and PWG raster otherwise. CUPS queues are reached over IPP at the
/// local scheduler; when the scheduler is not reachable that way, or the queue is a discovered destination that the scheduler has not created yet,
/// the CUPS command-line tools deliver the PDF instead. Windows queues receive raster sheets at the device resolution.
/// </remarks>
public sealed class PdfPrinterService : IPdfPrinterService {
    private readonly CupsScheduler _cups;

    /// <summary>Creates a service that uses the local CUPS scheduler where one is installed.</summary>
    public PdfPrinterService() : this(CupsScheduler.Local) { }

    internal PdfPrinterService(CupsScheduler cups) => _cups = cups;

    /// <inheritdoc />
    public Task<PdfPrinterCapabilities> GetCapabilitiesAsync(string printerName, CancellationToken cancellationToken = default) {
        ValidatePrinterName(printerName, nameof(printerName));
        cancellationToken.ThrowIfCancellationRequested();
        if (PdfPrinterAddress.TryGetIppUri(printerName, nameof(printerName), out Uri? uri))
            return IppPdfPrinter.GetCapabilitiesAsync(IppPrinterEndpoint.ForUri(printerName, uri!), cancellationToken);
        return OfficeRuntimeCompatibility.IsWindows
            ? WindowsPdfPrinter.GetCapabilitiesAsync(printerName, cancellationToken)
            : CupsPdfPrinter.GetCapabilitiesAsync(_cups, printerName, cancellationToken);
    }

    /// <inheritdoc />
    public Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CancellationToken cancellationToken = default) =>
        OfficeRuntimeCompatibility.IsWindows
            ? WindowsPdfPrinter.GetPrintersAsync(cancellationToken)
            : CupsPdfPrinter.GetPrintersAsync(_cups, cancellationToken);

    /// <inheritdoc />
    public Task<PdfPrintSubmission> SubmitAsync(PdfPrintJob job, PdfPrintDeliveryOptions options, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(job);
        ArgumentNullException.ThrowIfNull(options);
        PdfPrintDeliveryOptions snapshot = options.Snapshot();
        cancellationToken.ThrowIfCancellationRequested();
        if (PdfPrinterAddress.TryGetIppUri(snapshot.PrinterName, nameof(options), out Uri? uri))
            return IppPdfPrinter.SubmitAsync(IppPrinterEndpoint.ForUri(snapshot.PrinterName, uri!), job, snapshot, cancellationToken);
        return OfficeRuntimeCompatibility.IsWindows
            ? WindowsPdfPrinter.SubmitAsync(job, snapshot, cancellationToken)
            : CupsPdfPrinter.SubmitAsync(_cups, job, snapshot, cancellationToken);
    }

    private static void ValidatePrinterName(string printerName, string parameter) {
        if (printerName == null || string.IsNullOrWhiteSpace(printerName)) throw new ArgumentException("A non-empty value is required.", parameter);
        if (printerName.IndexOf('\0') >= 0) throw new ArgumentException("The printer name is invalid.", parameter);
    }
}

/// <summary>Recognizes printer URIs among printer names.</summary>
internal static class PdfPrinterAddress {
    /// <summary>Returns true for an ipp:// or ipps:// name; such a name must be a valid printer URI.</summary>
    internal static bool TryGetIppUri(string printerName, string parameter, out Uri? uri) {
        uri = null;
        if (!printerName.StartsWith("ipp://", StringComparison.OrdinalIgnoreCase) && !printerName.StartsWith("ipps://", StringComparison.OrdinalIgnoreCase)) return false;
        if (!Uri.TryCreate(printerName, UriKind.Absolute, out Uri? parsed) || string.IsNullOrEmpty(parsed.Host) ||
            !string.IsNullOrEmpty(parsed.UserInfo) || !string.IsNullOrEmpty(parsed.Fragment) || parsed.AbsoluteUri.Length > 1023)
            throw new ArgumentException("A printer URI needs a host and cannot contain credentials or a fragment.", parameter);
        uri = parsed;
        return true;
    }
}
