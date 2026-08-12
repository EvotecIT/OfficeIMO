using System.Collections.Generic;

namespace OfficeIMO.ChartForgeX;

/// <summary>Describes fidelity decisions made while converting one visual artifact.</summary>
public sealed class OfficeVisualConversionReport {
    private readonly List<string> _warnings = new List<string>();

    /// <summary>Gets whether the OfficeDrawing result retains vector content.</summary>
    public bool IsVector { get; internal set; }

    /// <summary>Gets whether the adapter used ChartForgeX PNG output instead of the imported SVG scene.</summary>
    public bool UsedRasterFallback { get; internal set; }

    /// <summary>Gets the number of SVG features the Office drawing importer could not represent completely.</summary>
    public int UnsupportedSvgFeatureCount { get; internal set; }

    /// <summary>Gets human-readable fidelity warnings.</summary>
    public IReadOnlyList<string> Warnings => _warnings;

    internal void Warn(string message) => _warnings.Add(message);
}
