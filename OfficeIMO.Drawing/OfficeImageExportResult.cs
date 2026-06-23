using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Result returned by dependency-free image export operations.
/// </summary>
public sealed class OfficeImageExportResult {
    /// <summary>
    /// Creates an image export result.
    /// </summary>
    public OfficeImageExportResult(
        OfficeImageExportFormat format,
        int width,
        int height,
        byte[] bytes,
        string? name = null,
        string? source = null,
        IReadOnlyList<OfficeImageExportDiagnostic>? diagnostics = null) {
        Format = format;
        Width = width;
        Height = height;
        Bytes = bytes ?? System.Array.Empty<byte>();
        Name = name;
        Source = source;
        Diagnostics = diagnostics ?? System.Array.Empty<OfficeImageExportDiagnostic>();
    }

    /// <summary>Output image format.</summary>
    public OfficeImageExportFormat Format { get; }

    /// <summary>Output width in pixels for PNG or CSS pixels for SVG.</summary>
    public int Width { get; }

    /// <summary>Output height in pixels for PNG or CSS pixels for SVG.</summary>
    public int Height { get; }

    /// <summary>Encoded image bytes.</summary>
    public byte[] Bytes { get; }

    /// <summary>Optional result name, such as a sheet or page name.</summary>
    public string? Name { get; }

    /// <summary>Optional source reference, such as a sheet range.</summary>
    public string? Source { get; }

    /// <summary>Diagnostics emitted while exporting.</summary>
    public IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }
}
