using System.Text;
using global::CodeGlyphX;
using global::CodeGlyphX.Rendering.Svg;

namespace OfficeIMO.Drawing.CodeGlyphX;

/// <summary>A symbol rendered once at its final physical size for every OfficeIMO format adapter.</summary>
internal sealed class OfficeBarcodeRendering {
    private readonly int _maximumElements;

    internal OfficeBarcodeRendering(string svg, int maximumElements, OfficeLength width, OfficeLength height, string alternativeText, string fileName) {
        Svg = Encoding.UTF8.GetBytes(svg);
        _maximumElements = maximumElements;
        Width = width;
        Height = height;
        AlternativeText = alternativeText;
        FileName = fileName;
    }

    /// <summary>UTF-8 SVG whose aspect ratio matches <see cref="Width"/> and <see cref="Height"/> exactly.</summary>
    internal byte[] Svg { get; }

    internal OfficeLength Width { get; }

    internal OfficeLength Height { get; }

    internal string AlternativeText { get; }

    internal string FileName { get; }

    /// <summary>Imports the SVG as a vector drawing measured in points at the placement size.</summary>
    internal OfficeDrawing ToDrawing() {
        OfficeDrawing source = CodeGlyphDrawingExtensions.ReadSvg(Encoding.UTF8.GetString(Svg), _maximumElements, out int unsupported);
        if (unsupported != 0) throw new NotSupportedException("The code appearance uses SVG features that cannot be kept as vector drawing content.");
        double width = Width.ToPoints(), height = Height.ToPoints();
        return new OfficeDrawing(width, height).AddEffectDrawing(source, OfficeTransform.Scale(width / source.Width, height / source.Height));
    }
}

/// <summary>Encodes a payload with CodeGlyphX and renders it at a physical size for document placement.</summary>
internal static class OfficeBarcodeRenderer {
    private static readonly OfficeLength DefaultMatrixWidth = OfficeLength.Inches(1.25D);
    private static readonly OfficeLength DefaultModuleWidth = OfficeLength.Millimeters(0.33D);
    private const int DefaultMatrixQuietZone = 4;
    private const int DefaultLinearQuietZone = 10;
    private const int DefaultBarHeightModules = 50;
    private const int LabelFontModules = 8;
    private const int LabelMarginModules = 2;
    private const string LabelFontFamily = "Arial, Helvetica, sans-serif";
    // Raster consumers (the Office PNG fallback and PDF conversion of Office pictures) draw the SVG from its intrinsic
    // pixel size, so whole-pixel modules at print resolution keep every module edge sharp.
    private const double SvgPixelsPerInch = 300D;

    internal static OfficeBarcodeRendering Render(SymbolFormat format, string payload, Action<OfficeBarcodeOptions>? configure) {
        if (payload == null) throw new ArgumentNullException(nameof(payload));
        var options = new OfficeBarcodeOptions();
        configure?.Invoke(options);
        Validate(format, options);

        SymbolCapability capability = SymbolCapabilities.Get(format);
        string alternativeText = options.AlternativeText ?? capability.DisplayName + ": " + payload;
        if (format == SymbolFormat.QrCode) {
            var encoding = new QrEncodingOptions();
            if (options.ErrorCorrection.HasValue) encoding.ErrorCorrectionLevel = options.ErrorCorrection.Value;
            return RenderMatrix(QR.Encode(payload, encoding).Modules, matrixSizing: true, options, alternativeText, "qr-code.svg");
        }

        if (capability.Family == SymbolFamily.Linear) {
            Barcode1D? linear = null;
            try {
                linear = Barcode.Encode(format, payload);
            } catch (NotSupportedException) {
                // Single-row symbols such as GS1 DataBar are encoded as module matrices.
            }
            if (linear != null) return RenderLinear(linear, capability, payload, options, alternativeText);
        }

        if (options.HumanReadableText == true)
            throw new ArgumentException($"{capability.DisplayName} has no human-readable text line; it applies to linear barcodes.", nameof(configure));
        MatrixSymbol symbol = MatrixBarcode.Encode(format, payload);
        return RenderMatrix(symbol.Modules, capability.Family == SymbolFamily.Matrix, options, alternativeText, "barcode.svg");
    }

    private static OfficeBarcodeRendering RenderMatrix(BitMatrix modules, bool matrixSizing, OfficeBarcodeOptions options, string alternativeText, string fileName) {
        int quietZone = options.QuietZone ?? DefaultMatrixQuietZone;
        int outWidth = modules.Width + quietZone * 2, outHeight = modules.Height + quietZone * 2;
        double modulePoints = ResolveModulePoints(outWidth, outHeight, matrixSizing, options);
        int pixelsPerModule = ResolvePixelsPerModule(modulePoints);
        string svg = MatrixSvgRenderer.Render(modules, new MatrixSvgRenderOptions {
            ModuleSize = pixelsPerModule,
            QuietZone = quietZone,
            DarkColor = ToCss(options.Color),
            LightColor = ToCss(options.BackgroundColor)
        });
        return new OfficeBarcodeRendering(svg,
            CodeGlyphDrawingExtensions.ResolveMaximumElements((long)modules.Width * modules.Height),
            OfficeLength.Points(outWidth * modulePoints), OfficeLength.Points(outHeight * modulePoints), alternativeText, fileName);
    }

    private static double ResolveModulePoints(int outWidth, int outHeight, bool matrixSizing, OfficeBarcodeOptions options) {
        if (options.Width.HasValue && options.Height.HasValue)
            return Math.Min(options.Width.Value.ToPoints() / outWidth, options.Height.Value.ToPoints() / outHeight);
        if (options.Width.HasValue) return options.Width.Value.ToPoints() / outWidth;
        if (options.Height.HasValue) return options.Height.Value.ToPoints() / outHeight;
        return matrixSizing ? DefaultMatrixWidth.ToPoints() / outWidth : DefaultModuleWidth.ToPoints();
    }

    private static OfficeBarcodeRendering RenderLinear(Barcode1D barcode, SymbolCapability capability, string payload, OfficeBarcodeOptions options, string alternativeText) {
        int quietZone = options.QuietZone ?? DefaultLinearQuietZone;
        int outWidth = barcode.TotalModules + quietZone * 2;
        string? label = options.HumanReadableText ?? true ? ResolveLabel(barcode, capability, payload) : null;
        int labelModules = string.IsNullOrEmpty(label) ? 0 : LabelFontModules + LabelMarginModules;

        double modulePoints = options.Width?.ToPoints() / outWidth ?? DefaultModuleWidth.ToPoints();
        int barHeightModules = DefaultBarHeightModules;
        if (options.Height.HasValue) {
            barHeightModules = (int)Math.Round(options.Height.Value.ToPoints() / modulePoints) - labelModules;
            if (barHeightModules < 1)
                throw new ArgumentOutOfRangeException(nameof(options), "The barcode height leaves no room for its bars; increase Height or reduce Width.");
        }

        int pixelsPerModule = ResolvePixelsPerModule(modulePoints);
        string svg = SvgBarcodeRenderer.Render(barcode, new BarcodeSvgRenderOptions {
            ModuleSize = pixelsPerModule,
            QuietZone = quietZone,
            HeightModules = barHeightModules,
            BarColor = ToCss(options.Color),
            BackgroundColor = ToCss(options.BackgroundColor),
            LabelColor = ToCss(options.Color),
            LabelText = label,
            LabelFontSize = LabelFontModules * pixelsPerModule,
            LabelMargin = LabelMarginModules * pixelsPerModule,
            LabelFontFamily = LabelFontFamily
        });
        return new OfficeBarcodeRendering(svg, CodeGlyphDrawingExtensions.ResolveMaximumElements(barcode.Segments.Count),
            OfficeLength.Points(outWidth * modulePoints), OfficeLength.Points((barHeightModules + labelModules) * modulePoints),
            alternativeText, "barcode.svg");
    }

    private static string ResolveLabel(Barcode1D barcode, SymbolCapability capability, string payload) {
        if (!string.IsNullOrWhiteSpace(barcode.HumanReadableText)) return barcode.HumanReadableText!;
        // Retail symbols print their check digit, which only the encoded bars carry, so read the label back from them.
        if (capability.Format is SymbolFormat.Ean or SymbolFormat.UpcA or SymbolFormat.UpcE) {
            var modules = new List<bool>(barcode.TotalModules + DefaultLinearQuietZone * 2);
            modules.AddRange(Enumerable.Repeat(false, DefaultLinearQuietZone));
            foreach (BarSegment segment in barcode.Segments) modules.AddRange(Enumerable.Repeat(segment.IsBar, segment.Modules));
            modules.AddRange(Enumerable.Repeat(false, DefaultLinearQuietZone));
            if (BarcodeDecoder.TryDecode(modules.ToArray(), capability.LegacyBarcodeType, out BarcodeDecoded decoded) &&
                !string.IsNullOrWhiteSpace(decoded.Text)) return decoded.Text;
        }
        return payload;
    }

    private static int ResolvePixelsPerModule(double modulePoints) {
        double pixels = modulePoints / 72D * SvgPixelsPerInch;
        return Math.Max(1, Math.Min(64, (int)Math.Ceiling(pixels)));
    }

    private static void Validate(SymbolFormat format, OfficeBarcodeOptions options) {
        if (options.Width.HasValue && options.Width.Value <= OfficeLength.Zero)
            throw new ArgumentOutOfRangeException(nameof(options.Width), "Width must be positive.");
        if (options.Height.HasValue && options.Height.Value <= OfficeLength.Zero)
            throw new ArgumentOutOfRangeException(nameof(options.Height), "Height must be positive.");
        if (options.QuietZone is < 0)
            throw new ArgumentOutOfRangeException(nameof(options.QuietZone), "The quiet zone cannot be negative.");
        if (options.ErrorCorrection.HasValue && format != SymbolFormat.QrCode)
            throw new ArgumentException("ErrorCorrection applies to QR codes only.", nameof(options.ErrorCorrection));
        RequireOpaqueLiteral(options.Color, nameof(options.Color));
        RequireOpaqueLiteral(options.BackgroundColor, nameof(options.BackgroundColor));
    }

    private static void RequireOpaqueLiteral(OfficeColor color, string name) {
        if (color.Role != OfficeColorRole.None || color.A != 255)
            throw new ArgumentException("Use a literal, opaque color so scanners see the same contrast in every format.", name);
    }

    private static string ToCss(OfficeColor color) => "#" + color.ToRgbHex();
}
