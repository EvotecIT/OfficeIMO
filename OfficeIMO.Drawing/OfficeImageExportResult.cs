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
        if (!System.Enum.IsDefined(typeof(OfficeImageExportFormat), format)) {
            throw new System.ArgumentOutOfRangeException(nameof(format));
        }
        if (width < 1) throw new System.ArgumentOutOfRangeException(nameof(width), "Image width must be positive.");
        if (height < 1) throw new System.ArgumentOutOfRangeException(nameof(height), "Image height must be positive.");
        if (bytes == null) throw new System.ArgumentNullException(nameof(bytes));
        if (!OfficeImageReader.TryIdentifyByContent(bytes, format.GetFileExtension(), out OfficeImageInfo identified) ||
            identified.Format != ToImageFormat(format)) {
            throw new System.ArgumentException(
                "Encoded image bytes do not match the declared " + format + " export format.",
                nameof(bytes));
        }
        if (identified.Width != width || identified.Height != height) {
            throw new System.ArgumentException(
                "Encoded image dimensions " + identified.Width + "x" + identified.Height +
                " do not match the declared " + width + "x" + height + " export dimensions.",
                nameof(bytes));
        }
        Format = format;
        Width = width;
        Height = height;
        _bytes = (byte[])bytes.Clone();
        Name = name;
        Source = source;
        Diagnostics = diagnostics == null
            ? System.Array.Empty<OfficeImageExportDiagnostic>()
            : new List<OfficeImageExportDiagnostic>(diagnostics).AsReadOnly();
    }

    /// <summary>Output image format.</summary>
    public OfficeImageExportFormat Format { get; }

    /// <summary>Output width in pixels for raster formats or CSS pixels for SVG.</summary>
    public int Width { get; }

    /// <summary>Output height in pixels for raster formats or CSS pixels for SVG.</summary>
    public int Height { get; }

    /// <summary>Canonical MIME type for the encoded output.</summary>
    public string MimeType => Format.GetMimeType();

    /// <summary>Canonical file extension, including the leading period.</summary>
    public string FileExtension => Format.GetFileExtension();

    /// <summary>Encoded image bytes.</summary>
    public byte[] Bytes => (byte[])_bytes.Clone();

    /// <summary>Optional result name, such as a sheet or page name.</summary>
    public string? Name { get; }

    /// <summary>Optional source reference, such as a sheet range.</summary>
    public string? Source { get; }

    /// <summary>Diagnostics emitted while exporting.</summary>
    public IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }

    private static OfficeImageFormat ToImageFormat(OfficeImageExportFormat format) => format switch {
        OfficeImageExportFormat.Png => OfficeImageFormat.Png,
        OfficeImageExportFormat.Svg => OfficeImageFormat.Svg,
        OfficeImageExportFormat.Jpeg => OfficeImageFormat.Jpeg,
        OfficeImageExportFormat.Tiff => OfficeImageFormat.Tiff,
        OfficeImageExportFormat.Webp => OfficeImageFormat.Webp,
        _ => OfficeImageFormat.Unknown
    };
}
