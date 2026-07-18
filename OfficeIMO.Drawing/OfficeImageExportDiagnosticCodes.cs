namespace OfficeIMO.Drawing;

/// <summary>Stable diagnostic codes shared by OfficeIMO image exporters.</summary>
public static class OfficeImageExportDiagnosticCodes {
    /// <summary>The requested raster scale was reduced to satisfy allocation or encoder limits.</summary>
    public const string RasterScaleReduced = "IMAGE_RASTER_SCALE_REDUCED";

    /// <summary>An embedded image used the caller-supplied codec.</summary>
    public const string SourceImageDecodedByCallerCodec = "IMAGE_SOURCE_DECODED_BY_CALLER_CODEC";

    /// <summary>An embedded image could not be decoded and was represented by a visible fallback.</summary>
    public const string SourceImageDecodeFallback = "IMAGE_SOURCE_DECODE_FALLBACK";

    /// <summary>A requested font face was unavailable and a deterministic fallback was selected.</summary>
    public const string FontSubstituted = "IMAGE_FONT_SUBSTITUTED";
}
