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

/// <summary>Delivery options applied to already reviewed print sheets.</summary>
public sealed class PdfPrintDeliveryOptions {
    /// <summary>Installed queue name.</summary>
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

    internal PdfPrintDeliveryOptions Snapshot() {
        ArgumentException.ThrowIfNullOrWhiteSpace(PrinterName);
        ArgumentException.ThrowIfNullOrWhiteSpace(DocumentName);
        if (PrinterName.IndexOf('\0') >= 0 || DocumentName.IndexOf('\0') >= 0) throw new ArgumentException("Printer and document names cannot contain null characters.");
        if (Copies is < 1 or > 100) throw new ArgumentOutOfRangeException(nameof(Copies));
        if (!Enum.IsDefined(Duplex)) throw new ArgumentOutOfRangeException(nameof(Duplex));
        if (PaperSourceId is not null && (string.IsNullOrWhiteSpace(PaperSourceId) || PaperSourceId.Length > 256 || PaperSourceId.Any(char.IsControl)))
            throw new ArgumentException("The paper-source identifier is invalid.", nameof(PaperSourceId));
        return new() { PrinterName = PrinterName, DocumentName = DocumentName, Copies = Copies, Duplex = Duplex, PaperSourceId = PaperSourceId, OutputFilePath = OutputFilePath };
    }
}

/// <summary>Receipt proving that a job was accepted by an operating-system print queue.</summary>
/// <remarks>Acceptance does not prove that paper or a virtual-printer file has finished printing.</remarks>
public sealed record PdfPrintSubmission(string PrinterName, string JobId, int SheetCount, int Copies, string? OutputFilePath) {
    /// <summary>Local staging cleanup problem after acceptance; the receipt remains valid and the job must not be resubmitted.</summary>
    public string? CleanupWarning { get; init; }
}

/// <summary>A delivery failure after submission began; retrying may print duplicate pages.</summary>
public sealed class PdfPrintDeliveryException : IOException {
    /// <summary>Creates an uncertain-delivery result for a printer provider after submission started.</summary>
    public PdfPrintDeliveryException(string? jobId, Exception inner) : base(
        "Printer delivery was interrupted after submission began. Check the printer queue before retrying; some pages may already have printed.", inner) => JobId = jobId;
    /// <summary>Known spooler job identifier, or null if the acknowledgement was interrupted.</summary>
    public string? JobId { get; }
}

/// <summary>Operating-system boundary for printer discovery and exact prepared-sheet delivery.</summary>
public interface IPdfPrinterService {
    /// <summary>Lists installed queues without changing printer settings.</summary>
    Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CancellationToken cancellationToken = default);
    /// <summary>Lists the selected queue's paper sources without changing printer settings.</summary>
    Task<IReadOnlyList<PdfPaperSourceInfo>> GetPaperSourcesAsync(string printerName, CancellationToken cancellationToken = default);
    /// <summary>Submits reviewed sheets and returns an acceptance receipt.</summary>
    Task<PdfPrintSubmission> SubmitAsync(PdfPreparedPrintDocument document, PdfPrintDeliveryOptions options, CancellationToken cancellationToken = default);
}

/// <summary>Windows GDI and macOS/Linux CUPS printer delivery.</summary>
public sealed class PdfPrinterService : IPdfPrinterService {
    /// <inheritdoc />
    public Task<IReadOnlyList<PdfPaperSourceInfo>> GetPaperSourcesAsync(string printerName, CancellationToken cancellationToken = default) {
        ArgumentException.ThrowIfNullOrWhiteSpace(printerName);
        if (printerName.Contains('\0')) throw new ArgumentException("The printer name is invalid.", nameof(printerName));
        cancellationToken.ThrowIfCancellationRequested();
        return OperatingSystem.IsWindows()
            ? WindowsPdfPrinter.GetPaperSourcesAsync(printerName, cancellationToken)
            : CupsPdfPrinter.GetPaperSourcesAsync(printerName, cancellationToken);
    }

    /// <inheritdoc />
    public Task<IReadOnlyList<PdfPrinterInfo>> GetPrintersAsync(CancellationToken cancellationToken = default) =>
        OperatingSystem.IsWindows()
            ? WindowsPdfPrinter.GetPrintersAsync(cancellationToken)
            : CupsPdfPrinter.GetPrintersAsync(cancellationToken);

    /// <inheritdoc />
    public Task<PdfPrintSubmission> SubmitAsync(PdfPreparedPrintDocument document, PdfPrintDeliveryOptions options, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(options);
        PdfPrintDeliveryOptions snapshot = options.Snapshot();
        if (document.Sheets.Count == 0) throw new ArgumentException("There are no prepared print sheets.", nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        return OperatingSystem.IsWindows()
            ? WindowsPdfPrinter.SubmitAsync(document, snapshot, cancellationToken)
            : CupsPdfPrinter.SubmitAsync(document, snapshot, cancellationToken);
    }
}
