using global::CodeGlyphX;
using OfficeIMO.Drawing.CodeGlyphX;

namespace OfficeIMO.Pdf;

/// <summary>Adds CodeGlyphX QR codes and barcodes to PDF content as vector drawings.</summary>
public static class PdfBarcodeExtensions {
    /// <summary>Adds a QR code at the current flow position.</summary>
    /// <param name="content">Content flow that receives the code.</param>
    /// <param name="payload">Text, URL or other data to encode.</param>
    /// <param name="configure">Optional size, colors, error correction and alternative text.</param>
    /// <param name="align">Horizontal alignment in the flow; left when omitted.</param>
    /// <returns>The same content builder, for further composition.</returns>
    public static PdfContentBuilder AddQrCode(this PdfContentBuilder content, string payload, Action<OfficeBarcodeOptions>? configure = null,
        OfficeAlignment? align = null) =>
        content.AddBarcode(SymbolFormat.QrCode, payload, configure, align);

    /// <summary>Adds a barcode of any CodeGlyphX symbology at the current flow position.</summary>
    /// <param name="content">Content flow that receives the code.</param>
    /// <param name="format">Symbology, such as <see cref="SymbolFormat.Code128"/>, <see cref="SymbolFormat.Ean"/> or <see cref="SymbolFormat.DataMatrix"/>.</param>
    /// <param name="payload">Data to encode; the symbology decides which characters are valid.</param>
    /// <param name="configure">Optional size, colors, quiet zone, human-readable text and alternative text.</param>
    /// <param name="align">Horizontal alignment in the flow; left when omitted.</param>
    /// <returns>The same content builder, for further composition.</returns>
    public static PdfContentBuilder AddBarcode(this PdfContentBuilder content, SymbolFormat format, string payload,
        Action<OfficeBarcodeOptions>? configure = null, OfficeAlignment? align = null) {
        if (content == null) throw new ArgumentNullException(nameof(content));
        OfficeBarcodeRendering rendering = OfficeBarcodeRenderer.Render(format, payload, configure);
        return content.Drawing(rendering.ToDrawing(), align, style: new PdfDrawingStyle { AlternativeText = rendering.AlternativeText });
    }

    /// <summary>Draws a QR code with its top-left corner at a page position.</summary>
    /// <param name="canvas">Page canvas that receives the code.</param>
    /// <param name="payload">Text, URL or other data to encode.</param>
    /// <param name="x">Distance from the page's left edge.</param>
    /// <param name="y">Distance from the page's top edge.</param>
    /// <param name="configure">Optional size, colors, error correction and alternative text.</param>
    /// <returns>The same canvas, for further drawing.</returns>
    public static PdfPageCanvas QrCode(this PdfPageCanvas canvas, string payload, OfficeLength x, OfficeLength y,
        Action<OfficeBarcodeOptions>? configure = null) =>
        canvas.Barcode(SymbolFormat.QrCode, payload, x, y, configure);

    /// <summary>Draws a barcode of any CodeGlyphX symbology with its top-left corner at a page position.</summary>
    /// <param name="canvas">Page canvas that receives the code.</param>
    /// <param name="format">Symbology, such as <see cref="SymbolFormat.Code128"/>, <see cref="SymbolFormat.Ean"/> or <see cref="SymbolFormat.DataMatrix"/>.</param>
    /// <param name="payload">Data to encode; the symbology decides which characters are valid.</param>
    /// <param name="x">Distance from the page's left edge.</param>
    /// <param name="y">Distance from the page's top edge.</param>
    /// <param name="configure">Optional size, colors, quiet zone, human-readable text and alternative text.</param>
    /// <returns>The same canvas, for further drawing.</returns>
    public static PdfPageCanvas Barcode(this PdfPageCanvas canvas, SymbolFormat format, string payload, OfficeLength x, OfficeLength y,
        Action<OfficeBarcodeOptions>? configure = null) {
        if (canvas == null) throw new ArgumentNullException(nameof(canvas));
        OfficeBarcodeRendering rendering = OfficeBarcodeRenderer.Render(format, payload, configure);
        return canvas.Drawing(rendering.ToDrawing(), x, y, rendering.Width, rendering.Height,
            new PdfDrawingStyle { AlternativeText = rendering.AlternativeText });
    }
}
