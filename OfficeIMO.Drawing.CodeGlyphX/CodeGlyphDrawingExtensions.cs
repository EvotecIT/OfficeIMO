using System;
using System.Text;
using global::CodeGlyphX;
using global::CodeGlyphX.Rendering.Svg;

namespace OfficeIMO.Drawing.CodeGlyphX;

/// <summary>
/// Converts typed CodeGlyphX symbols to OfficeIMO.Drawing scenes through the neutral SVG boundary.
/// </summary>
public static class CodeGlyphDrawingExtensions {
    /// <summary>Renders a QR code and imports it as an OfficeIMO drawing.</summary>
    public static OfficeDrawing ToOfficeDrawing(this QrCode qrCode, QrSvgRenderOptions? options = null) =>
        ToOfficeDrawing(qrCode, out _, options);

    /// <summary>Renders a QR code and imports it while reporting SVG features that required fallback.</summary>
    public static OfficeDrawing ToOfficeDrawing(
        this QrCode qrCode,
        out int unsupportedFeatureCount,
        QrSvgRenderOptions? options = null) {
        if (qrCode is null) throw new ArgumentNullException(nameof(qrCode));
        string svg = SvgQrRenderer.Render(qrCode.Modules, options ?? new QrSvgRenderOptions());
        return ReadSvg(svg, out unsupportedFeatureCount);
    }

    /// <summary>Renders a generic matrix symbol and imports it as an OfficeIMO drawing.</summary>
    public static OfficeDrawing ToOfficeDrawing(this BitMatrix modules, MatrixSvgRenderOptions? options = null) =>
        ToOfficeDrawing(modules, out _, options);

    /// <summary>Renders a generic matrix symbol and imports it while reporting SVG features that required fallback.</summary>
    public static OfficeDrawing ToOfficeDrawing(
        this BitMatrix modules,
        out int unsupportedFeatureCount,
        MatrixSvgRenderOptions? options = null) {
        if (modules is null) throw new ArgumentNullException(nameof(modules));
        string svg = MatrixSvgRenderer.Render(modules, options ?? new MatrixSvgRenderOptions());
        return ReadSvg(svg, out unsupportedFeatureCount);
    }

    /// <summary>Renders a linear barcode and imports it as an OfficeIMO drawing.</summary>
    public static OfficeDrawing ToOfficeDrawing(this Barcode1D barcode, BarcodeSvgRenderOptions? options = null) =>
        ToOfficeDrawing(barcode, out _, options);

    /// <summary>Renders a linear barcode and imports it while reporting SVG features that required fallback.</summary>
    public static OfficeDrawing ToOfficeDrawing(
        this Barcode1D barcode,
        out int unsupportedFeatureCount,
        BarcodeSvgRenderOptions? options = null) {
        if (barcode is null) throw new ArgumentNullException(nameof(barcode));
        string svg = SvgBarcodeRenderer.Render(barcode, options ?? new BarcodeSvgRenderOptions());
        return ReadSvg(svg, out unsupportedFeatureCount);
    }

    private static OfficeDrawing ReadSvg(string svg, out int unsupportedFeatureCount) {
        byte[] bytes = Encoding.UTF8.GetBytes(svg);
        if (!OfficeSvgDrawingReader.TryRead(bytes, out OfficeDrawing? drawing, out unsupportedFeatureCount) || drawing is null) {
            throw new InvalidOperationException("CodeGlyphX produced SVG that OfficeIMO.Drawing could not read.");
        }
        return drawing;
    }
}
