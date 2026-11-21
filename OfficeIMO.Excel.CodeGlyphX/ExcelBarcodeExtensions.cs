using global::CodeGlyphX;
using OfficeIMO.Drawing.CodeGlyphX;

namespace OfficeIMO.Excel;

/// <summary>Adds CodeGlyphX QR codes and barcodes to worksheets as SVG pictures with a PNG fallback.</summary>
public static class ExcelBarcodeExtensions {
    /// <summary>Adds a QR code anchored at a cell and returns the picture.</summary>
    /// <param name="sheet">Worksheet that receives the picture.</param>
    /// <param name="cell">A1 reference of the anchor cell, such as <c>B2</c>.</param>
    /// <param name="payload">Text, URL or other data to encode.</param>
    /// <param name="configure">Optional size, colors, error correction and alternative text.</param>
    /// <returns>The native picture, for offsets, placement and other worksheet settings.</returns>
    public static ExcelImage AddQrCode(this ExcelSheet sheet, string cell, string payload, Action<OfficeBarcodeOptions>? configure = null) =>
        sheet.AddBarcode(cell, SymbolFormat.QrCode, payload, configure);

    /// <summary>Adds a QR code anchored at a one-based row and column and returns the picture.</summary>
    /// <param name="sheet">Worksheet that receives the picture.</param>
    /// <param name="row">One-based anchor row.</param>
    /// <param name="column">One-based anchor column.</param>
    /// <param name="payload">Text, URL or other data to encode.</param>
    /// <param name="configure">Optional size, colors, error correction and alternative text.</param>
    /// <returns>The native picture, for offsets, placement and other worksheet settings.</returns>
    public static ExcelImage AddQrCode(this ExcelSheet sheet, int row, int column, string payload, Action<OfficeBarcodeOptions>? configure = null) =>
        sheet.AddBarcode(row, column, SymbolFormat.QrCode, payload, configure);

    /// <summary>Adds a barcode of any CodeGlyphX symbology anchored at a cell and returns the picture.</summary>
    /// <param name="sheet">Worksheet that receives the picture.</param>
    /// <param name="cell">A1 reference of the anchor cell, such as <c>B2</c>.</param>
    /// <param name="format">Symbology, such as <see cref="SymbolFormat.Code128"/>, <see cref="SymbolFormat.Ean"/> or <see cref="SymbolFormat.DataMatrix"/>.</param>
    /// <param name="payload">Data to encode; the symbology decides which characters are valid.</param>
    /// <param name="configure">Optional size, colors, quiet zone, human-readable text and alternative text.</param>
    /// <returns>The native picture, for offsets, placement and other worksheet settings.</returns>
    public static ExcelImage AddBarcode(this ExcelSheet sheet, string cell, SymbolFormat format, string payload, Action<OfficeBarcodeOptions>? configure = null) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        OfficeBarcodeRendering rendering = OfficeBarcodeRenderer.Render(format, payload, configure);
        return sheet.AddImage(cell, rendering.Svg, CreateImageOptions(rendering));
    }

    /// <summary>Adds a barcode of any CodeGlyphX symbology anchored at a one-based row and column and returns the picture.</summary>
    /// <param name="sheet">Worksheet that receives the picture.</param>
    /// <param name="row">One-based anchor row.</param>
    /// <param name="column">One-based anchor column.</param>
    /// <param name="format">Symbology, such as <see cref="SymbolFormat.Code128"/>, <see cref="SymbolFormat.Ean"/> or <see cref="SymbolFormat.DataMatrix"/>.</param>
    /// <param name="payload">Data to encode; the symbology decides which characters are valid.</param>
    /// <param name="configure">Optional size, colors, quiet zone, human-readable text and alternative text.</param>
    /// <returns>The native picture, for offsets, placement and other worksheet settings.</returns>
    public static ExcelImage AddBarcode(this ExcelSheet sheet, int row, int column, SymbolFormat format, string payload, Action<OfficeBarcodeOptions>? configure = null) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        OfficeBarcodeRendering rendering = OfficeBarcodeRenderer.Render(format, payload, configure);
        return sheet.AddImage(row, column, rendering.Svg, CreateImageOptions(rendering));
    }

    private static ExcelImageOptions CreateImageOptions(OfficeBarcodeRendering rendering) => new ExcelImageOptions {
        ContentType = "image/svg+xml",
        Width = rendering.Width,
        Height = rendering.Height,
        Description = rendering.AlternativeText,
        LockAspectRatio = true
    };
}
