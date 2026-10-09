using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal static class VisioDrawingTextAlignment {
    internal static OfficeTextAlignment ToOfficeTextAlignment(VisioTextHorizontalAlignment? alignment) {
        switch (alignment) {
            case VisioTextHorizontalAlignment.Left:
                return OfficeTextAlignment.Left;
            case VisioTextHorizontalAlignment.Right:
                return OfficeTextAlignment.Right;
            case VisioTextHorizontalAlignment.Justify:
                return OfficeTextAlignment.Justify;
            default:
                return OfficeTextAlignment.Center;
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
