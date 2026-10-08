using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

/// <summary>Immutable ODF line-leader declarations for an explicit paragraph tab.</summary>
public sealed class OdfTabLineLeader {
    /// <summary>Creates a native line declaration with a pattern, line count, width and optional RGB color.</summary>
    /// <param name="style">none, solid, dotted, dash, long-dash, dot-dash, dot-dot-dash or wave.</param>
    /// <param name="type">none, single or double.</param>
    /// <param name="width">auto, normal, thin, medium, bold, thick, a positive integer, percentage or absolute ODF length.</param>
    /// <param name="color">Explicit RGB color; null writes font-color.</param>
    /// <remarks>Shared projection uses an explicit approximation profile for named, integer and percentage widths. Absolute lengths retain their point width. A textual leader takes precedence.</remarks>
    public OdfTabLineLeader(string style = "solid", string type = "single", string width = "auto", OdfColor? color = null) {
        Style = style; Type = type; Width = width; Color = color;
        _ = ToDrawingLeader();
    }
    /// <summary>Native line-pattern token.</summary>
    public string Style { get; }
    /// <summary>Native line-count token.</summary>
    public string Type { get; }
    /// <summary>Native width token retained without normalization.</summary>
    public string Width { get; }
    /// <summary>Explicit RGB color, or null for font-color.</summary>
    public OdfColor? Color { get; }

    internal OfficeTextTabLineLeader ToDrawingLeader() {
        OfficeTextTabLineLeaderStyle style = Style switch {
            "none" => OfficeTextTabLineLeaderStyle.None, "solid" => OfficeTextTabLineLeaderStyle.Solid,
            "dotted" => OfficeTextTabLineLeaderStyle.Dotted, "dash" => OfficeTextTabLineLeaderStyle.Dash,
            "long-dash" => OfficeTextTabLineLeaderStyle.LongDash, "dot-dash" => OfficeTextTabLineLeaderStyle.DotDash,
            "dot-dot-dash" => OfficeTextTabLineLeaderStyle.DotDotDash, "wave" => OfficeTextTabLineLeaderStyle.Wave,
            _ => throw new ArgumentException("Unknown ODF leader line style.", nameof(Style))
        };
        if (Type is not ("none" or "single" or "double")) throw new ArgumentException("Unknown ODF leader line type.", nameof(Type));
        double fraction = Width switch {
            "auto" or "normal" or "medium" => .05D, "thin" => .025D, "bold" or "thick" => .1D, _ => 0D
        };
        double? points = null;
        if (fraction == 0) {
            if (Width != null && Width.EndsWith("%", StringComparison.Ordinal) &&
                double.TryParse(Width.Substring(0, Width.Length - 1), NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                    CultureInfo.InvariantCulture, out double percentage)) fraction = percentage * .0005D;
            else if (double.TryParse(Width, NumberStyles.Integer, CultureInfo.InvariantCulture, out double multiple)) fraction = multiple * .05D;
            else if (Width != null && OdfLength.Parse(Width).TryToPoints(out double length)) { points = length; fraction = .05D; }
            else throw new ArgumentException("Unknown ODF leader width.", nameof(Width));
        }
        // The shared constructor enforces positive finite widths before any geometry allocation.
        return new OfficeTextTabLineLeader(Type == "none" ? OfficeTextTabLineLeaderStyle.None : style,
            Type == "double", points, fraction, Color.HasValue ? OfficeColor.Parse(Color.Value.ToString()) : null);
    }
}
