using global::CodeGlyphX;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.CodeGlyphX;

namespace OfficeIMO.PowerPoint;

/// <summary>Adds CodeGlyphX QR codes and barcodes to slides as SVG pictures with a PNG fallback.</summary>
public static class PowerPointBarcodeExtensions {
    /// <summary>Adds a QR code with its top-left corner at a slide position and returns the picture.</summary>
    /// <param name="slide">Slide that receives the picture.</param>
    /// <param name="payload">Text, URL or other data to encode.</param>
    /// <param name="left">Distance from the slide's left edge.</param>
    /// <param name="top">Distance from the slide's top edge.</param>
    /// <param name="configure">Optional size, colors, error correction and alternative text.</param>
    /// <returns>The native picture, for animation, links and other slide settings.</returns>
    public static PowerPointImage AddQrCode(this PowerPointSlide slide, string payload, OfficeLength left, OfficeLength top,
        Action<OfficeBarcodeOptions>? configure = null) =>
        slide.AddBarcode(SymbolFormat.QrCode, payload, left, top, configure);

    /// <summary>Adds a QR code as large as fits in a layout box, centered in it, and returns the picture.</summary>
    /// <param name="slide">Slide that receives the picture.</param>
    /// <param name="payload">Text, URL or other data to encode.</param>
    /// <param name="box">Area the code fits into, such as a cell of <c>SlideSize.GetContentBox(...)</c>.</param>
    /// <param name="configure">Optional colors, error correction and alternative text; a width or height set here replaces the box size.</param>
    /// <returns>The native picture, for animation, links and other slide settings.</returns>
    public static PowerPointImage AddQrCode(this PowerPointSlide slide, string payload, PowerPointLayoutBox box,
        Action<OfficeBarcodeOptions>? configure = null) =>
        slide.AddBarcode(SymbolFormat.QrCode, payload, box, configure);

    /// <summary>Adds a barcode of any CodeGlyphX symbology with its top-left corner at a slide position and returns the picture.</summary>
    /// <param name="slide">Slide that receives the picture.</param>
    /// <param name="format">Symbology, such as <see cref="SymbolFormat.Code128"/>, <see cref="SymbolFormat.Ean"/> or <see cref="SymbolFormat.DataMatrix"/>.</param>
    /// <param name="payload">Data to encode; the symbology decides which characters are valid.</param>
    /// <param name="left">Distance from the slide's left edge.</param>
    /// <param name="top">Distance from the slide's top edge.</param>
    /// <param name="configure">Optional size, colors, quiet zone, human-readable text and alternative text.</param>
    /// <returns>The native picture, for animation, links and other slide settings.</returns>
    public static PowerPointImage AddBarcode(this PowerPointSlide slide, SymbolFormat format, string payload, OfficeLength left, OfficeLength top,
        Action<OfficeBarcodeOptions>? configure = null) {
        if (slide == null) throw new ArgumentNullException(nameof(slide));
        return Place(slide, OfficeBarcodeRenderer.Render(format, payload, configure), left, top);
    }

    /// <summary>Adds a barcode of any CodeGlyphX symbology as large as fits in a layout box, centered in it, and returns the picture.</summary>
    /// <param name="slide">Slide that receives the picture.</param>
    /// <param name="format">Symbology, such as <see cref="SymbolFormat.Code128"/>, <see cref="SymbolFormat.Ean"/> or <see cref="SymbolFormat.DataMatrix"/>.</param>
    /// <param name="payload">Data to encode; the symbology decides which characters are valid.</param>
    /// <param name="box">Area the code fits into; linear barcodes fill its width and height.</param>
    /// <param name="configure">Optional colors, quiet zone, human-readable text and alternative text; a width or height set here replaces the box size.</param>
    /// <returns>The native picture, for animation, links and other slide settings.</returns>
    public static PowerPointImage AddBarcode(this PowerPointSlide slide, SymbolFormat format, string payload, PowerPointLayoutBox box,
        Action<OfficeBarcodeOptions>? configure = null) {
        if (slide == null) throw new ArgumentNullException(nameof(slide));
        OfficeBarcodeRendering rendering = OfficeBarcodeRenderer.Render(format, payload, options => {
            options.Width = box.Width;
            options.Height = box.Height;
            configure?.Invoke(options);
        });
        OfficeLength left = box.Left + OfficeLength.Emus((box.Width.ToEmus() - rendering.Width.ToEmus()) / 2);
        OfficeLength top = box.Top + OfficeLength.Emus((box.Height.ToEmus() - rendering.Height.ToEmus()) / 2);
        return Place(slide, rendering, left, top);
    }

    private static PowerPointImage Place(PowerPointSlide slide, OfficeBarcodeRendering rendering, OfficeLength left, OfficeLength top) {
        using var stream = new MemoryStream(rendering.Svg, writable: false);
        PowerPointImage picture = slide.AddImage(stream, OfficeImageFormat.Svg, left, top, rendering.Width, rendering.Height);
        picture.AltText = rendering.AlternativeText;
        return picture;
    }
}
