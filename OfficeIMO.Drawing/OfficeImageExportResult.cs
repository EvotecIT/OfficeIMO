using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing.Internal;

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
        IReadOnlyList<OfficeImageExportDiagnostic>? diagnostics = null,
        string? savedPath = null) {
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
        DpiX = identified.DpiX;
        DpiY = identified.DpiY;
        _bytes = (byte[])bytes.Clone();
        Name = name;
        Source = source;
        SavedPath = string.IsNullOrWhiteSpace(savedPath) ? null : Path.GetFullPath(savedPath);
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

    /// <summary>Horizontal encoded resolution in dots per inch, defaulting to 96 when absent.</summary>
    public double DpiX { get; }

    /// <summary>Vertical encoded resolution in dots per inch, defaulting to 96 when absent.</summary>
    public double DpiY { get; }

    /// <summary>Physical width in inches according to encoded dimensions and resolution.</summary>
    public double PhysicalWidthInches => Width / DpiX;

    /// <summary>Physical height in inches according to encoded dimensions and resolution.</summary>
    public double PhysicalHeightInches => Height / DpiY;

    /// <summary>Canonical MIME type for the encoded output.</summary>
    public string MimeType => Format.GetMimeType();

    /// <summary>Canonical file extension, including the leading period.</summary>
    public string FileExtension => Format.GetFileExtension();

    /// <summary>Encoded image bytes.</summary>
    public byte[] Bytes => (byte[])_bytes.Clone();

    /// <summary>Encoded byte count without copying the payload.</summary>
    public long EncodedLength => _bytes.LongLength;

    /// <summary>Optional result name, such as a sheet or page name.</summary>
    public string? Name { get; }

    /// <summary>Optional source reference, such as a sheet range.</summary>
    public string? Source { get; }

    /// <summary>Normalized file path when this result was committed by a save operation.</summary>
    public string? SavedPath { get; }

    /// <summary>Diagnostics emitted while exporting.</summary>
    public IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }

    /// <summary>Creates an aggregate report for this result.</summary>
    public OfficeImageExportReport CreateReport() => new OfficeImageExportReport(new[] { this });

    /// <summary>Throws when this result violates the supplied acceptance policy.</summary>
    public OfficeImageExportResult Require(OfficeImageExportPolicy policy) {
        if (policy == null) throw new System.ArgumentNullException(nameof(policy));
        policy.EnsureAccepted(Diagnostics);
        return this;
    }

    /// <summary>
    /// Saves the encoded image to a file, appending the canonical extension when the path has none.
    /// </summary>
    public OfficeImageExportResult Save(
        string path,
        OfficeImageExportFileConflictPolicy conflictPolicy = OfficeImageExportFileConflictPolicy.FailIfExists) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new System.ArgumentException("Output path cannot be null or whitespace.", nameof(path));
        }

        string resolvedPath = OfficeImageExportPath.ResolveFile(path, Format, conflictPolicy);
        OfficeFileCommit.WriteAllBytes(
            resolvedPath,
            _bytes,
            OfficeImageExportPath.ToCommitPolicy(conflictPolicy));
        return WithSavedPath(resolvedPath);
    }

    /// <summary>Writes the encoded image to a caller-owned stream without closing it.</summary>
    public OfficeImageExportResult Save(Stream stream) {
        OfficeStreamWriter.WriteAllBytes(stream, _bytes);
        return this;
    }

    /// <summary>
    /// Asynchronously saves the encoded image to a file, appending the canonical extension when the path has none.
    /// </summary>
    public async Task<OfficeImageExportResult> SaveAsync(
        string path,
        OfficeImageExportFileConflictPolicy conflictPolicy = OfficeImageExportFileConflictPolicy.FailIfExists,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new System.ArgumentException("Output path cannot be null or whitespace.", nameof(path));
        }

        cancellationToken.ThrowIfCancellationRequested();
        string resolvedPath = OfficeImageExportPath.ResolveFile(path, Format, conflictPolicy);
        await OfficeFileCommit.WriteAllBytesAsync(
            resolvedPath,
            _bytes,
            OfficeImageExportPath.ToCommitPolicy(conflictPolicy),
            cancellationToken).ConfigureAwait(false);
        return WithSavedPath(resolvedPath);
    }

    /// <summary>Asynchronously writes the encoded image to a caller-owned stream without closing it.</summary>
    public async Task<OfficeImageExportResult> SaveAsync(
        Stream stream,
        CancellationToken cancellationToken = default) {
        await OfficeStreamWriter.WriteAllBytesAsync(stream, _bytes, cancellationToken).ConfigureAwait(false);
        return this;
    }

    internal OfficeImageExportResult WithSavedPath(string path) => new OfficeImageExportResult(
        Format,
        Width,
        Height,
        _bytes,
        Name,
        Source,
        Diagnostics,
        path);

    private static OfficeImageFormat ToImageFormat(OfficeImageExportFormat format) => format switch {
        OfficeImageExportFormat.Png => OfficeImageFormat.Png,
        OfficeImageExportFormat.Svg => OfficeImageFormat.Svg,
        OfficeImageExportFormat.Jpeg => OfficeImageFormat.Jpeg,
        OfficeImageExportFormat.Tiff => OfficeImageFormat.Tiff,
        OfficeImageExportFormat.Webp => OfficeImageFormat.Webp,
        _ => OfficeImageFormat.Unknown
    };
}
