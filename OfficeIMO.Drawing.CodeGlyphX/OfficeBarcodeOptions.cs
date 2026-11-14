using global::CodeGlyphX;

namespace OfficeIMO.Drawing.CodeGlyphX;

/// <summary>
/// Size, appearance and accessibility of a QR code or barcode placed by the OfficeIMO CodeGlyphX adapters.
/// Every property is optional; the defaults produce a scannable symbol at a sensible document size.
/// </summary>
public sealed class OfficeBarcodeOptions {
    /// <summary>
    /// Width of the whole symbol, including its quiet zone. Two-dimensional codes default to 1.25 inches; linear,
    /// stacked and postal codes default to 0.33 mm per module. With only a width, the symbol keeps its proportions.
    /// </summary>
    public OfficeLength? Width { get; set; }

    /// <summary>
    /// Height of the whole symbol, including any human-readable text. Linear barcodes adjust their bar height to it;
    /// other symbols keep their proportions and fit inside <see cref="Width"/> and <see cref="Height"/>.
    /// </summary>
    public OfficeLength? Height { get; set; }

    /// <summary>Light margin around the symbol in modules. Defaults to 4 for two-dimensional codes and 10 for linear barcodes.</summary>
    public int? QuietZone { get; set; }

    /// <summary>Color of the bars or dark modules. Use a literal, opaque color.</summary>
    public OfficeColor Color { get; set; } = OfficeColor.Black;

    /// <summary>Color of the background and quiet zone. Use a literal, opaque color.</summary>
    public OfficeColor BackgroundColor { get; set; } = OfficeColor.White;

    /// <summary>QR error correction level. Applies to QR codes only; CodeGlyphX chooses the level when this is not set.</summary>
    public QrErrorCorrectionLevel? ErrorCorrection { get; set; }

    /// <summary>
    /// Whether a linear barcode prints its human-readable text under the bars. Linear barcodes show it by default;
    /// other symbologies do not support it.
    /// </summary>
    public bool? HumanReadableText { get; set; }

    /// <summary>
    /// Accessible description of the symbol. Defaults to the symbology and its payload, for example
    /// <c>QR Code: https://evotec.xyz</c>, so screen readers can announce what the code contains.
    /// </summary>
    public string? AlternativeText { get; set; }
}
