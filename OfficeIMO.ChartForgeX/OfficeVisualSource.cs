using System;
using System.Text;

namespace OfficeIMO.ChartForgeX;

/// <summary>Describes a portable SVG visual for Office conversion without requiring a live ChartForgeX model.</summary>
public sealed class OfficeVisualSource {
    private readonly byte[] _svgBytes;
    private string _id = "visual";
    private string _title = string.Empty;
    private string _alternativeText = string.Empty;

    /// <summary>Initializes a portable SVG source.</summary>
    public OfficeVisualSource(byte[] svgBytes) {
        if (svgBytes == null) throw new ArgumentNullException(nameof(svgBytes));
        if (svgBytes.Length == 0) throw new ArgumentException("SVG payload cannot be empty.", nameof(svgBytes));
        _svgBytes = (byte[])svgBytes.Clone();
    }

    /// <summary>Initializes a portable SVG source from markup.</summary>
    public OfficeVisualSource(string svg) : this(Encoding.UTF8.GetBytes(
        string.IsNullOrWhiteSpace(svg) ? throw new ArgumentException("SVG markup cannot be empty.", nameof(svg)) : svg)) {
    }

    /// <summary>Gets or sets a stable source identifier.</summary>
    public string Id { get => _id; set => _id = string.IsNullOrWhiteSpace(value) ? throw new ArgumentException("Visual id cannot be empty.", nameof(value)) : value.Trim(); }

    /// <summary>Gets or sets a human-friendly title.</summary>
    public string Title { get => _title; set => _title = value ?? throw new ArgumentNullException(nameof(value)); }

    /// <summary>Gets or sets the accessible description used by Office placements.</summary>
    public string AlternativeText { get => _alternativeText; set => _alternativeText = value ?? throw new ArgumentNullException(nameof(value)); }

    /// <summary>Returns an independent copy of the SVG payload.</summary>
    public byte[] GetSvgBytes() => (byte[])_svgBytes.Clone();
}
