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
    public OfficeVisualSource(string svg) : this(EncodeSvgMarkupAsUtf8(svg)) {
    }

    /// <summary>Gets or sets a stable source identifier.</summary>
    public string Id { get => _id; set => _id = string.IsNullOrWhiteSpace(value) ? throw new ArgumentException("Visual id cannot be empty.", nameof(value)) : value.Trim(); }

    /// <summary>Gets or sets a human-friendly title.</summary>
    public string Title { get => _title; set => _title = value ?? throw new ArgumentNullException(nameof(value)); }

    /// <summary>Gets or sets the accessible description used by Office placements.</summary>
    public string AlternativeText { get => _alternativeText; set => _alternativeText = value ?? throw new ArgumentNullException(nameof(value)); }

    /// <summary>Gets or sets whether the visual is decorative and should be omitted from accessibility structure.</summary>
    public bool IsDecorative { get; set; }

    /// <summary>Returns an independent copy of the SVG payload.</summary>
    public byte[] GetSvgBytes() => (byte[])_svgBytes.Clone();

    private static byte[] EncodeSvgMarkupAsUtf8(string svg) {
        if (string.IsNullOrWhiteSpace(svg)) {
            throw new ArgumentException("SVG markup cannot be empty.", nameof(svg));
        }

        string normalized = svg[0] == '\uFEFF' ? svg.Substring(1) : svg;
        if (normalized.StartsWith("<?xml", StringComparison.Ordinal)) {
            int declarationEnd = normalized.IndexOf("?>", StringComparison.Ordinal);
            if (declarationEnd < 0) {
                throw new ArgumentException("SVG XML declaration is incomplete.", nameof(svg));
            }

            normalized = normalized.Substring(declarationEnd + 2);
            if (string.IsNullOrWhiteSpace(normalized)) {
                throw new ArgumentException("SVG markup cannot be empty.", nameof(svg));
            }
        }

        return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(normalized);
    }
}
