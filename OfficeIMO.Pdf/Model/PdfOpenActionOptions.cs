namespace OfficeIMO.Pdf;

/// <summary>
/// Defines the initial destination a PDF viewer should open when displaying a generated document.
/// </summary>
public sealed class PdfOpenActionOptions {
    /// <summary>Creates a generated document open action targeting a one-based page number.</summary>
    public PdfOpenActionOptions(
        int pageNumber = 1,
        double? destinationTop = null,
        PdfOpenActionDestinationMode destinationMode = PdfOpenActionDestinationMode.Xyz,
        double? destinationLeft = null,
        double? destinationBottom = null,
        double? destinationRight = null) {
        if (pageNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), "PDF open-action page number must be positive.");
        }

        ValidateCoordinate(destinationTop, nameof(destinationTop));
        ValidateCoordinate(destinationLeft, nameof(destinationLeft));
        ValidateCoordinate(destinationBottom, nameof(destinationBottom));
        ValidateCoordinate(destinationRight, nameof(destinationRight));

        if (destinationMode != PdfOpenActionDestinationMode.Xyz &&
            destinationMode != PdfOpenActionDestinationMode.Fit &&
            destinationMode != PdfOpenActionDestinationMode.FitHorizontal &&
            destinationMode != PdfOpenActionDestinationMode.FitVertical &&
            destinationMode != PdfOpenActionDestinationMode.FitRectangle &&
            destinationMode != PdfOpenActionDestinationMode.FitBoundingBox &&
            destinationMode != PdfOpenActionDestinationMode.FitBoundingBoxHorizontal &&
            destinationMode != PdfOpenActionDestinationMode.FitBoundingBoxVertical) {
            throw new ArgumentOutOfRangeException(nameof(destinationMode), destinationMode, "PDF open-action destination mode is not supported.");
        }

        PageNumber = pageNumber;
        DestinationTop = destinationTop;
        DestinationLeft = destinationLeft;
        DestinationBottom = destinationBottom;
        DestinationRight = destinationRight;
        DestinationMode = destinationMode;
    }

    /// <summary>One-based generated page number to show when the document opens.</summary>
    public int PageNumber { get; }

    /// <summary>Optional top coordinate for the destination. When omitted, the generated page top is used.</summary>
    public double? DestinationTop { get; }

    /// <summary>Optional left coordinate for destination modes that use a left or rectangle coordinate.</summary>
    public double? DestinationLeft { get; }

    /// <summary>Optional bottom coordinate for rectangle destinations.</summary>
    public double? DestinationBottom { get; }

    /// <summary>Optional right coordinate for rectangle destinations.</summary>
    public double? DestinationRight { get; }

    /// <summary>Viewer destination mode emitted for the open action.</summary>
    public PdfOpenActionDestinationMode DestinationMode { get; }

    internal PdfOpenActionOptions Clone() => new PdfOpenActionOptions(PageNumber, DestinationTop, DestinationMode, DestinationLeft, DestinationBottom, DestinationRight);

    private static void ValidateCoordinate(double? value, string parameterName) {
        if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value))) {
            throw new ArgumentOutOfRangeException(parameterName, "PDF open-action destination coordinate must be finite.");
        }
    }
}
