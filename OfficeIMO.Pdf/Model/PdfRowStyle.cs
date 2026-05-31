namespace OfficeIMO.Pdf;

/// <summary>
/// Describes reusable layout rhythm for row/column blocks.
/// </summary>
public sealed class PdfRowStyle {
    /// <summary>Built-in gutter used for multi-column rows when neither the row nor a default row style specifies one.</summary>
    public const double DefaultGap = 18D;

    private double? _gap;
    private double _spacingBefore;
    private double _spacingAfter;
    private double _columnSeparatorWidth;

    /// <summary>Horizontal gutter between row columns, in points. When null the row uses its explicit gap or the built-in Word-like gutter.</summary>
    public double? Gap {
        get => _gap;
        set {
            ValidateOptionalNonNegativeFiniteValue(value, nameof(Gap), "Row gap must be a non-negative finite value.");
            _gap = value;
        }
    }

    /// <summary>Vertical space before the row, in points.</summary>
    public double SpacingBefore {
        get => _spacingBefore;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingBefore), "Row spacing before must be a non-negative finite value.");
            _spacingBefore = value;
        }
    }

    /// <summary>Vertical space after the row, in points.</summary>
    public double SpacingAfter {
        get => _spacingAfter;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingAfter), "Row spacing after must be a non-negative finite value.");
            _spacingAfter = value;
        }
    }

    /// <summary>Optional vertical separator color drawn between row columns.</summary>
    public PdfColor? ColumnSeparatorColor { get; set; }

    /// <summary>Vertical separator line width, in points. A non-positive width disables separator drawing.</summary>
    public double ColumnSeparatorWidth {
        get => _columnSeparatorWidth;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(ColumnSeparatorWidth), "Row column separator width must be a non-negative finite value.");
            _columnSeparatorWidth = value;
        }
    }

    /// <summary>When true, the row moves to the next page instead of splitting across pages.</summary>
    public bool KeepTogether { get; set; }

    /// <summary>When true, the row moves with the first visible part of the following block when they fit together.</summary>
    public bool KeepWithNext { get; set; }

    /// <summary>Creates a copy of this row style.</summary>
    public PdfRowStyle Clone() {
        return new PdfRowStyle {
            Gap = Gap,
            SpacingBefore = SpacingBefore,
            SpacingAfter = SpacingAfter,
            ColumnSeparatorColor = ColumnSeparatorColor,
            ColumnSeparatorWidth = ColumnSeparatorWidth,
            KeepTogether = KeepTogether,
            KeepWithNext = KeepWithNext
        };
    }

    private static void ValidateNonNegativeFiniteValue(double value, string paramName, string message) {
        if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentException(message, paramName);
        }
    }

    private static void ValidateOptionalNonNegativeFiniteValue(double? value, string paramName, string message) {
        if (value.HasValue) {
            ValidateNonNegativeFiniteValue(value.Value, paramName, message);
        }
    }
}
