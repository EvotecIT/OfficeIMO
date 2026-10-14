using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

/// <summary>Reference edge or corner for an explicit glue point.</summary>
public enum OdgGluePointAlignment {
    /// <summary>Center of the shape.</summary>
    Center,
    /// <summary>Top-left corner.</summary>
    TopLeft,
    /// <summary>Top edge center.</summary>
    Top,
    /// <summary>Top-right corner.</summary>
    TopRight,
    /// <summary>Left edge center.</summary>
    Left,
    /// <summary>Right edge center.</summary>
    Right,
    /// <summary>Bottom-left corner.</summary>
    BottomLeft,
    /// <summary>Bottom-right corner.</summary>
    BottomRight,
    /// <summary>Bottom edge center.</summary>
    Bottom
}

/// <summary>Directions in which native routing may leave a glue point.</summary>
public enum OdgGluePointEscapeDirection {
    /// <summary>Let the native router choose.</summary>
    Auto,
    /// <summary>Leave toward the left.</summary>
    Left,
    /// <summary>Leave toward the right.</summary>
    Right,
    /// <summary>Leave upward.</summary>
    Up,
    /// <summary>Leave downward.</summary>
    Down,
    /// <summary>Leave horizontally in either direction.</summary>
    Horizontal,
    /// <summary>Leave vertically in either direction.</summary>
    Vertical
}

/// <summary>A persistent attachment point on a Draw shape.</summary>
public sealed class OdgGluePoint {
    internal OdgShape Shape { get; }
    internal XElement Element { get; }
    internal OdgGluePoint(OdgShape shape, XElement element) { Shape = shape; Element = element; }
    /// <summary>Native glue point identifier, unique within its shape.</summary>
    public int Id => int.Parse((string?)Element.Attribute(OdfNamespaces.Draw + "id") ?? throw new InvalidDataException("Missing glue point ID."), CultureInfo.InvariantCulture);
    /// <summary>
    /// Native departure constraint. Editing a shared point affects every connector referencing it;
    /// saved path projection does not recalculate routes from this constraint.
    /// </summary>
    public OdgGluePointEscapeDirection EscapeDirection {
        get => ((string?)Element.Attribute(OdfNamespaces.Draw + "escape-direction")) switch {
            null or "auto" => OdgGluePointEscapeDirection.Auto,
            "left" => OdgGluePointEscapeDirection.Left, "right" => OdgGluePointEscapeDirection.Right,
            "up" => OdgGluePointEscapeDirection.Up, "down" => OdgGluePointEscapeDirection.Down,
            "horizontal" => OdgGluePointEscapeDirection.Horizontal, "vertical" => OdgGluePointEscapeDirection.Vertical,
            _ => throw new InvalidDataException("Invalid glue point escape direction.")
        };
        set {
            if (!Enum.IsDefined(typeof(OdgGluePointEscapeDirection), value)) throw new ArgumentOutOfRangeException(nameof(value));
            Element.SetAttributeValue(OdfNamespaces.Draw + "escape-direction", value.ToString().ToLowerInvariant()); Shape.Dirty();
        }
    }
    internal OfficePoint Position {
        get {
            OdfRect bounds = Shape.AttachmentBounds;
            double w = bounds.Width.ToPoints(), h = bounds.Height.ToPoints();
            string? alignment = (string?)Element.Attribute(OdfNamespaces.Draw + "align");
            double x = ReadOffset("x", w, alignment), y = ReadOffset("y", h, alignment);
            (double ax, double ay) = alignment switch {
                null or "center" => (0.5, 0.5), "top-left" => (0, 0), "top" => (0.5, 0), "top-right" => (1, 0),
                "left" => (0, 0.5), "right" => (1, 0.5), "bottom-left" => (0, 1), "bottom" => (0.5, 1), "bottom-right" => (1, 1),
                _ => throw new NotSupportedException("Unsupported glue point alignment.")
            };
            return Shape.PageTransform.TransformPoint(new OfficePoint(bounds.X.ToPoints() + ax * w + x, bounds.Y.ToPoints() + ay * h + y));
        }
    }
    private double ReadOffset(string name, double dimension, string? alignment) {
        string value = (string?)Element.Attribute(OdfNamespaces.Svg + name) ?? "0%";
        if (alignment == null && value.EndsWith("%", StringComparison.Ordinal))
            return double.Parse(value.Substring(0, value.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture) * dimension / 100;
        double distance = OdfLength.Parse(value).ToPoints();
        // Older Draw producers encode relative offsets as 1/100 mm lengths over a 10,000-unit canvas.
        // An explicit alignment selects actual absolute offsets; without it these are relative.
        return alignment == null ? dimension * distance * 2540 / 72 / 10000 : distance;
    }
}
