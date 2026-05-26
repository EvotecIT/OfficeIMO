namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable horizontal rule appearance and rhythm style.
/// </summary>
public sealed class PdfHorizontalRuleStyle {
    private double _thickness = 0.5;
    private double _spacingBefore = 6;
    private double _spacingAfter = 6;

    /// <summary>Rule thickness in points.</summary>
    public double Thickness {
        get => _thickness;
        set {
            ValidatePositiveFiniteValue(value, nameof(Thickness), "Horizontal rule thickness must be a positive finite value.");
            _thickness = value;
        }
    }

    /// <summary>Rule stroke color.</summary>
    public PdfColor Color { get; set; } = PdfColor.Gray;

    /// <summary>Vertical space before the rule in the surrounding document flow, in points.</summary>
    public double SpacingBefore {
        get => _spacingBefore;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingBefore), "Horizontal rule spacing before must be a non-negative finite value.");
            _spacingBefore = value;
        }
    }

    /// <summary>Vertical space after the rule in the surrounding document flow, in points.</summary>
    public double SpacingAfter {
        get => _spacingAfter;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingAfter), "Horizontal rule spacing after must be a non-negative finite value.");
            _spacingAfter = value;
        }
    }

    /// <summary>When true, the rule moves to the next page when it would otherwise be separated from the following flow block.</summary>
    public bool KeepWithNext { get; set; }

    /// <summary>Creates a copy of this horizontal rule style.</summary>
    public PdfHorizontalRuleStyle Clone() {
        return new PdfHorizontalRuleStyle {
            Thickness = Thickness,
            Color = Color,
            SpacingBefore = SpacingBefore,
            SpacingAfter = SpacingAfter,
            KeepWithNext = KeepWithNext
        };
    }

    private static void ValidatePositiveFiniteValue(double value, string paramName, string message) {
        if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentException(message, paramName);
        }
    }

    private static void ValidateNonNegativeFiniteValue(double value, string paramName, string message) {
        if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentException(message, paramName);
        }
    }
}
