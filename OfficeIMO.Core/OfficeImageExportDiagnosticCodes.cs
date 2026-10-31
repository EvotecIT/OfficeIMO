namespace OfficeIMO.Drawing;

/// <summary>Stable diagnostic codes shared by OfficeIMO image exporters.</summary>
internal static class OfficeImageExportDiagnosticCodes {
    /// <summary>The requested raster scale was reduced to satisfy allocation or encoder limits.</summary>
    public const string RasterScaleReduced = global::OfficeIMO.OfficeDiagnosticCodes.ImageRasterScaleReduced;

    /// <summary>An embedded image used the caller-supplied codec.</summary>
    public const string SourceImageDecodedByCallerCodec = global::OfficeIMO.OfficeDiagnosticCodes.ImageSourceDecodedByCallerCodec;

    /// <summary>An embedded image could not be decoded and was represented by a visible fallback.</summary>
    public const string SourceImageDecodeFallback = global::OfficeIMO.OfficeDiagnosticCodes.ImageSourceDecodeFallback;

    /// <summary>An embedded image could not be decoded and was omitted from raster output.</summary>
    public const string SourceImageDecodeOmitted = global::OfficeIMO.OfficeDiagnosticCodes.ImageSourceDecodeOmitted;

    /// <summary>A static embedded image discards animation, additional frames, or pages.</summary>
    public const string SourceImageStaticFrameSelected = global::OfficeIMO.OfficeDiagnosticCodes.ImageSourceStaticFrameSelected;

    /// <summary>An embedded SVG preview exceeded structural budgets or used unsupported visual features.</summary>
    public const string SourceSvgPreviewLoss = global::OfficeIMO.OfficeDiagnosticCodes.ImageSourceSvgPreviewLoss;

    /// <summary>A requested font face was unavailable and a deterministic fallback was selected.</summary>
    public const string FontSubstituted = global::OfficeIMO.OfficeDiagnosticCodes.FontFamilySubstitution;

    /// <summary>Complex text used the bounded managed fallback because no complete host shaping result was available.</summary>
    public const string TextShapingFallback = global::OfficeIMO.OfficeDiagnosticCodes.ImageTextShapingFallback;
}
