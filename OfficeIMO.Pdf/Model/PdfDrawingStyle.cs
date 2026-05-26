namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable placement and rhythm style for OfficeIMO.Drawing-backed flow objects.
/// </summary>
public sealed class PdfDrawingStyle {
    private PdfAlign _align = PdfAlign.Left;
    private double _spacingBefore;
    private double _spacingAfter;

    /// <summary>Object alignment within the current content frame.</summary>
    public PdfAlign Align {
        get => _align;
        set {
            Guard.LeftCenterRightAlign(value, nameof(Align), "Drawing");
            _align = value;
        }
    }

    /// <summary>Vertical space before the drawing object in the surrounding document flow, in points.</summary>
    public double SpacingBefore {
        get => _spacingBefore;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingBefore), "Drawing spacing before must be a non-negative finite value.");
            _spacingBefore = value;
        }
    }

    /// <summary>Vertical space after the drawing object in the surrounding document flow, in points.</summary>
    public double SpacingAfter {
        get => _spacingAfter;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingAfter), "Drawing spacing after must be a non-negative finite value.");
            _spacingAfter = value;
        }
    }

    /// <summary>Moves the drawing object to the next page with the first visible part of the following block when they fit together.</summary>
    public bool KeepWithNext { get; set; }

    /// <summary>Creates a copy of this drawing placement style.</summary>
    public PdfDrawingStyle Clone() {
        return new PdfDrawingStyle {
            Align = Align,
            SpacingBefore = SpacingBefore,
            SpacingAfter = SpacingAfter,
            KeepWithNext = KeepWithNext
        };
    }

    private static void ValidateNonNegativeFiniteValue(double value, string paramName, string message) {
        if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentException(message, paramName);
        }
    }
}
