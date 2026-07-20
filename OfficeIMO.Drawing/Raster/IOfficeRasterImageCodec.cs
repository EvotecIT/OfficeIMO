namespace OfficeIMO.Drawing;

/// <summary>Optional image-codec boundary for formats not decoded by the dependency-free Drawing core.</summary>
public interface IOfficeRasterImageCodec {
    /// <summary>Attempts to decode encoded image bytes into an RGBA raster.</summary>
    bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image);
}

/// <summary>Options for rasterizing an Office drawing with an optional image codec.</summary>
public sealed class OfficeDrawingRasterRenderOptions {
    /// <summary>Output scale. Defaults to 1.</summary>
    public double Scale { get; set; } = 1D;
    /// <summary>Optional canvas background.</summary>
    public OfficeColor? Background { get; set; }
    /// <summary>Optional decoder for formats not handled by the dependency-free core.</summary>
    public IOfficeRasterImageCodec? ImageCodec { get; set; }
    /// <summary>Optional host text shaper used for complex-script glyph positioning.</summary>
    public IOfficeTextShapingProvider? TextShapingProvider { get; set; }
    /// <summary>Optional BCP 47 language hint passed to the text shaper.</summary>
    public string? TextShapingLanguage { get; set; }
    /// <summary>Optional sink for fidelity diagnostics discovered while rasterizing the drawing.</summary>
    public System.Collections.Generic.ICollection<OfficeImageExportDiagnostic>? DiagnosticSink { get; set; }
    /// <summary>Optional source label attached to rasterization diagnostics.</summary>
    public string? DiagnosticSource { get; set; }
    /// <summary>Cancellation observed between drawing elements and nested render stages.</summary>
    public System.Threading.CancellationToken CancellationToken { get; set; }
}
