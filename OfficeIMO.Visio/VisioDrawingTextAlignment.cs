using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal static class VisioDrawingTextAlignment {
    internal static OfficeTextAlignment ToOfficeTextAlignment(OfficeAlignment? alignment) {
        switch (alignment) {
            case OfficeAlignment.Left:
            case OfficeAlignment.Start:
                return OfficeTextAlignment.Left;
            case OfficeAlignment.Right:
            case OfficeAlignment.End:
                return OfficeTextAlignment.Right;
            case OfficeAlignment.Justify:
                return OfficeTextAlignment.Justify;
            default:
                return OfficeTextAlignment.Center;
        }
    }

    // Visio HorzAlign has no logical start/end values; shape text flows left to right, so Start and End write Left and Right.
    internal static int ToHorzAlignCell(OfficeAlignment alignment) => alignment switch {
        OfficeAlignment.Left or OfficeAlignment.Start => 0,
        OfficeAlignment.Center => 1,
        OfficeAlignment.Right or OfficeAlignment.End => 2,
        OfficeAlignment.Justify => 3,
        _ => throw new ArgumentOutOfRangeException(nameof(alignment), alignment, "Unknown Visio text alignment.")
    };

    // Force-justify (4) and other native values have no shared alignment and are not projected.
    internal static bool TryFromHorzAlignCell(int value, out OfficeAlignment alignment) {
        switch (value) {
            case 0: alignment = OfficeAlignment.Left; return true;
            case 1: alignment = OfficeAlignment.Center; return true;
            case 2: alignment = OfficeAlignment.Right; return true;
            case 3: alignment = OfficeAlignment.Justify; return true;
            default: alignment = OfficeAlignment.Left; return false;
        }
    }

    internal static OfficeTextVerticalAlignment ToOfficeTextVerticalAlignment(VisioTextVerticalAlignment? alignment) {
        switch (alignment) {
            case VisioTextVerticalAlignment.Top:
                return OfficeTextVerticalAlignment.Top;
            case VisioTextVerticalAlignment.Bottom:
                return OfficeTextVerticalAlignment.Bottom;
            default:
                return OfficeTextVerticalAlignment.Center;
        }
    }
}
