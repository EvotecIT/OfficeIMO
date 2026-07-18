using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>Metadata for an image result that was committed without retaining its payload.</summary>
public sealed class OfficeImageExportSavedFile {
    private readonly ReadOnlyCollection<OfficeImageExportDiagnostic> _diagnostics;

    internal OfficeImageExportSavedFile(OfficeImageExportResult result, string path) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        Path = System.IO.Path.GetFullPath(path);
        Format = result.Format;
        Width = result.Width;
        Height = result.Height;
        DpiX = result.DpiX;
        DpiY = result.DpiY;
        EncodedLength = result.EncodedLength;
        Name = result.Name;
        Source = result.Source;
        _diagnostics = Array.AsReadOnly(new List<OfficeImageExportDiagnostic>(result.Diagnostics).ToArray());
    }

    /// <summary>Normalized committed file path.</summary>
    public string Path { get; }

    /// <summary>Encoded image format.</summary>
    public OfficeImageExportFormat Format { get; }

    /// <summary>Encoded pixel or CSS-pixel width.</summary>
    public int Width { get; }

    /// <summary>Encoded pixel or CSS-pixel height.</summary>
    public int Height { get; }

    /// <summary>Horizontal encoded resolution.</summary>
    public double DpiX { get; }

    /// <summary>Vertical encoded resolution.</summary>
    public double DpiY { get; }

    /// <summary>Encoded byte count written to disk.</summary>
    public long EncodedLength { get; }

    /// <summary>Optional result name.</summary>
    public string? Name { get; }

    /// <summary>Optional source reference.</summary>
    public string? Source { get; }

    /// <summary>Diagnostics emitted while rendering this file.</summary>
    public IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics => _diagnostics;
}
