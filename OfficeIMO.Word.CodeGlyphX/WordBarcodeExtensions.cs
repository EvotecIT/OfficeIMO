using global::CodeGlyphX;
using OfficeIMO.Drawing.CodeGlyphX;

namespace OfficeIMO.Word;

/// <summary>Adds CodeGlyphX QR codes and barcodes to Word documents as SVG pictures with a PNG fallback.</summary>
public static class WordBarcodeExtensions {
    /// <summary>Adds a QR code to the paragraph and returns the picture.</summary>
    /// <param name="paragraph">Paragraph that receives the inline picture.</param>
    /// <param name="payload">Text, URL or other data to encode.</param>
    /// <param name="configure">Optional size, colors, error correction and alternative text.</param>
    /// <returns>The native picture, for wrapping, alignment and other Word settings.</returns>
    public static WordImage AddQrCode(this WordParagraph paragraph, string payload, Action<OfficeBarcodeOptions>? configure = null) =>
        paragraph.AddBarcode(SymbolFormat.QrCode, payload, configure);

    /// <summary>Adds a barcode of any CodeGlyphX symbology to the paragraph and returns the picture.</summary>
    /// <param name="paragraph">Paragraph that receives the inline picture.</param>
    /// <param name="format">Symbology, such as <see cref="SymbolFormat.Code128"/>, <see cref="SymbolFormat.Ean"/> or <see cref="SymbolFormat.DataMatrix"/>.</param>
    /// <param name="payload">Data to encode; the symbology decides which characters are valid.</param>
    /// <param name="configure">Optional size, colors, quiet zone, human-readable text and alternative text.</param>
    /// <returns>The native picture, for wrapping, alignment and other Word settings.</returns>
    public static WordImage AddBarcode(this WordParagraph paragraph, SymbolFormat format, string payload, Action<OfficeBarcodeOptions>? configure = null) {
        if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
        OfficeBarcodeRendering rendering = OfficeBarcodeRenderer.Render(format, payload, configure);
        using var stream = new MemoryStream(rendering.Svg, writable: false);
        return paragraph.AddImage(stream, rendering.FileName, rendering.Width, rendering.Height,
            new WordImageOptions { AltText = rendering.AlternativeText });
    }

    /// <summary>Adds a QR code in a new paragraph and returns the picture.</summary>
    /// <param name="document">Document that receives a new paragraph.</param>
    /// <param name="payload">Text, URL or other data to encode.</param>
    /// <param name="configure">Optional size, colors, error correction and alternative text.</param>
    /// <returns>The native picture, for wrapping, alignment and other Word settings.</returns>
    public static WordImage AddQrCode(this WordDocument document, string payload, Action<OfficeBarcodeOptions>? configure = null) =>
        document.AddBarcode(SymbolFormat.QrCode, payload, configure);

    /// <summary>Adds a barcode of any CodeGlyphX symbology in a new paragraph and returns the picture.</summary>
    /// <param name="document">Document that receives a new paragraph.</param>
    /// <param name="format">Symbology, such as <see cref="SymbolFormat.Code128"/>, <see cref="SymbolFormat.Ean"/> or <see cref="SymbolFormat.DataMatrix"/>.</param>
    /// <param name="payload">Data to encode; the symbology decides which characters are valid.</param>
    /// <param name="configure">Optional size, colors, quiet zone, human-readable text and alternative text.</param>
    /// <returns>The native picture, for wrapping, alignment and other Word settings.</returns>
    public static WordImage AddBarcode(this WordDocument document, SymbolFormat format, string payload, Action<OfficeBarcodeOptions>? configure = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        // Render before adding the paragraph so an invalid payload leaves the document unchanged.
        OfficeBarcodeRendering rendering = OfficeBarcodeRenderer.Render(format, payload, configure);
        using var stream = new MemoryStream(rendering.Svg, writable: false);
        return document.AddParagraph().AddImage(stream, rendering.FileName, rendering.Width, rendering.Height,
            new WordImageOptions { AltText = rendering.AlternativeText });
    }
}
