namespace OfficeIMO.Pdf;

/// <summary>
/// Describes reusable layout rhythm for row/column blocks.
/// </summary>
public sealed class PdfRowStyle {
    /// <summary>Built-in gutter used for multi-column rows when neither the row nor a default row style specifies one.</summary>
    public static readonly global::OfficeIMO.OfficeLength DefaultGap = global::OfficeIMO.OfficeLength.Points(18D);

    private global::OfficeIMO.OfficeLength? _gap;
    private global::OfficeIMO.OfficeLength _spacingBefore;
    private global::OfficeIMO.OfficeLength _spacingAfter;
    private global::OfficeIMO.OfficeLength _columnSeparatorWidth;

    /// <summary>Horizontal gutter between row columns, in points. When null the row uses its explicit gap or the built-in Word-like gutter.</summary>
    public global::OfficeIMO.OfficeLength? Gap {
        get => _gap;
        set {
            ValidateOptionalNonNegativeFiniteValue((value?.ToPoints()), nameof(Gap), "Row gap must be a non-negative finite value.");
            _gap = value;
        }
    }

    /// <summary>Vertical space before the row, in points.</summary>
    public global::OfficeIMO.OfficeLength SpacingBefore {
        get => _spacingBefore;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(SpacingBefore), "Row spacing before must be a non-negative finite value.");
            _spacingBefore = value;
        }
    }

    /// <summary>Vertical space after the row, in points.</summary>
    public global::OfficeIMO.OfficeLength SpacingAfter {
        get => _spacingAfter;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(SpacingAfter), "Row spacing after must be a non-negative finite value.");
            _spacingAfter = value;
        }
    }

    /// <summary>Optional vertical separator color drawn between row columns.</summary>
    public global::OfficeIMO.OfficeColor? ColumnSeparatorColor { get; set; }

    /// <summary>Vertical separator line width, in points. A non-positive width disables separator drawing.</summary>
    public global::OfficeIMO.OfficeLength ColumnSeparatorWidth {
        get => _columnSeparatorWidth;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(ColumnSeparatorWidth), "Row column separator width must be a non-negative finite value.");
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
