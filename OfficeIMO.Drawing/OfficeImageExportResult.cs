using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Result returned by dependency-free image export operations.
/// </summary>
public sealed class OfficeImageExportResult {
    private readonly byte[] _bytes;

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
        _bytes = bytes == null ? System.Array.Empty<byte>() : (byte[])bytes.Clone();
        Name = name;
        Source = source;
        Diagnostics = diagnostics == null
            ? System.Array.Empty<OfficeImageExportDiagnostic>()
            : new List<OfficeImageExportDiagnostic>(diagnostics).AsReadOnly();
    }

    /// <summary>Output image format.</summary>
    public OfficeImageExportFormat Format { get; }

    /// <summary>Output width in pixels for PNG or CSS pixels for SVG.</summary>
    public int Width { get; }

    /// <summary>Output height in pixels for PNG or CSS pixels for SVG.</summary>
    public int Height { get; }

    /// <summary>Encoded image bytes.</summary>
    public byte[] Bytes => (byte[])_bytes.Clone();

    /// <summary>Optional result name, such as a sheet or page name.</summary>
    public string? Name { get; }

    /// <summary>Optional source reference, such as a sheet range.</summary>
    public string? Source { get; }

    /// <summary>Diagnostics emitted while exporting.</summary>
    public IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }
}
