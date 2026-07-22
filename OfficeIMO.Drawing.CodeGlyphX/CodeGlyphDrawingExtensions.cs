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
    /// <exception cref="NotSupportedException">The SVG options include a raster logo that the drawing importer cannot preserve.</exception>
    public static OfficeDrawing ToOfficeDrawing(this QrCode qrCode, QrSvgRenderOptions? options = null) =>
        ToOfficeDrawing(qrCode, out _, options);

    /// <summary>Renders a QR code and imports it while reporting SVG features that required fallback.</summary>
    /// <exception cref="NotSupportedException">The SVG options include a raster logo that the drawing importer cannot preserve.</exception>
    public static OfficeDrawing ToOfficeDrawing(
        this QrCode qrCode,
        out int unsupportedFeatureCount,
        QrSvgRenderOptions? options = null) {
        if (qrCode is null) throw new ArgumentNullException(nameof(qrCode));
        QrSvgRenderOptions renderOptions = options ?? new QrSvgRenderOptions();
        if (renderOptions.Logo is not null) {
            throw new NotSupportedException("QR logo images cannot be represented by the OfficeIMO.Drawing SVG importer.");
        }
        BitMatrix modules = qrCode.Modules;
        string svg = SvgQrRenderer.Render(modules, renderOptions);
        int elementsPerCell = renderOptions.ModuleShape == global::CodeGlyphX.Rendering.Png.QrPngModuleShape.DotGrid ? 4 : 1;
        return ReadSvg(svg, ResolveMaximumElements(modules, elementsPerCell), out unsupportedFeatureCount);
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
        return ReadSvg(svg, ResolveMaximumElements(modules, elementsPerCell: 1), out unsupportedFeatureCount);
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
        return ReadSvg(svg, ResolveMaximumElements(barcode.Segments.Count), out unsupportedFeatureCount);
    }

    private static int ResolveMaximumElements(BitMatrix modules, int elementsPerCell) {
        return ResolveMaximumElements((long)modules.Width * modules.Height * elementsPerCell);
    }

    private static int ResolveMaximumElements(long symbolElements) {
        long requested = symbolElements + 1024L;
        return (int)Math.Min(
            OfficeSvgDrawingReaderOptions.MaximumAllowedElements,
            Math.Max(OfficeSvgDrawingReaderOptions.DefaultMaximumElements, requested));
    }

    private static OfficeDrawing ReadSvg(string svg, int maximumElements, out int unsupportedFeatureCount) {
        byte[] bytes = Encoding.UTF8.GetBytes(svg);
        var readerOptions = new OfficeSvgDrawingReaderOptions {
            MaximumElements = maximumElements,
            MaximumViewportDimension = OfficeSvgDrawingReaderOptions.MaximumAllowedViewportDimension,
            MaximumViewportPixels = OfficeSvgDrawingReaderOptions.MaximumAllowedViewportPixels
        };
        if (!OfficeSvgDrawingReader.TryRead(bytes, readerOptions, out OfficeDrawing? drawing, out unsupportedFeatureCount) || drawing is null) {
            throw new InvalidOperationException("CodeGlyphX produced SVG that OfficeIMO.Drawing could not read.");
        }
        return drawing;
    }
}
