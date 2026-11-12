using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>A paper size reported by a printer queue.</summary>
/// <param name="Id">Provider identifier: a Windows paper code or a CUPS media keyword.</param>
/// <param name="Name">Driver-supplied display name.</param>
/// <param name="Size">Portrait paper dimensions, or null when the provider does not state them.</param>
public sealed record PdfPrinterMedia(string Id, string Name, PdfPageSize? Size) {
    /// <inheritdoc />
    public override string ToString() => Name;
}

/// <summary>A printer resolution in dots per inch.</summary>
public sealed record PdfPrinterResolution(int HorizontalDpi, int VerticalDpi) {
    /// <inheritdoc />
    public override string ToString() => HorizontalDpi == VerticalDpi ? HorizontalDpi + " dpi" : HorizontalDpi + "x" + VerticalDpi + " dpi";
}

/// <summary>Media, trays and features a printer queue reports. Null values mean the queue does not say.</summary>
public sealed class PdfPrinterCapabilities {
    /// <summary>Queue name.</summary>
    public required string PrinterName { get; init; }
    /// <summary>Supported paper sizes in driver order.</summary>
    public IReadOnlyList<PdfPrinterMedia> Media { get; init; } = [];
    /// <summary>The queue's default paper, when reported.</summary>
    public PdfPrinterMedia? DefaultMedia { get; init; }
    /// <summary>Paper sources; pass an <see cref="PdfPaperSourceInfo.Id"/> as <see cref="PdfPrintDeliveryOptions.PaperSourceId"/>.</summary>
    public IReadOnlyList<PdfPaperSourceInfo> PaperSources { get; init; } = [];
    /// <summary>Whether two-sided printing is available.</summary>
    public bool? SupportsDuplex { get; init; }
    /// <summary>Whether the device prints in color.</summary>
    public bool? SupportsColor { get; init; }
    /// <summary>Whether the queue or driver produces collated copies itself.</summary>
    public bool? SupportsCollation { get; init; }
    /// <summary>Highest copy count the driver produces itself.</summary>
    public int? MaximumCopies { get; init; }
    /// <summary>Supported resolutions.</summary>
    public IReadOnlyList<PdfPrinterResolution> Resolutions { get; init; } = [];
}
