using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    private static int[] NormalizePageNumbers(int[]? pageNumbers, int pageCount) {
        if (pageNumbers is null || pageNumbers.Length == 0) {
            return Enumerable.Range(1, pageCount).ToArray();
        }

        var seen = new HashSet<int>();
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            if (pageNumber < 1 || pageNumber > pageCount) {
                throw new ArgumentOutOfRangeException(nameof(pageNumbers), "Page number " + pageNumber.ToString(CultureInfo.InvariantCulture) + " is outside the document page range 1-" + pageCount.ToString(CultureInfo.InvariantCulture) + ".");
            }

            if (!seen.Add(pageNumber)) {
                throw new ArgumentException("Duplicate page selections are not supported.", nameof(pageNumbers));
            }
        }

        return pageNumbers;
    }

    private static void ValidateOptions(PdfTextStampOptions options) {
        if (options.FontSize <= 0 || double.IsNaN(options.FontSize) || double.IsInfinity(options.FontSize)) {
            throw new ArgumentOutOfRangeException(nameof(options), "Font size must be a positive finite value.");
        }

        if ((options.X.HasValue && (double.IsNaN(options.X.Value) || double.IsInfinity(options.X.Value))) ||
            (options.Y.HasValue && (double.IsNaN(options.Y.Value) || double.IsInfinity(options.Y.Value))) ||
            double.IsNaN(options.RotationDegrees) ||
            double.IsInfinity(options.RotationDegrees)) {
            throw new ArgumentOutOfRangeException(nameof(options), "Text stamp coordinates and rotation must be finite.");
        }
    }
    private static void ValidateImageOptions(PdfImageStampOptions options) {
        if (options.Width.HasValue && (options.Width.Value <= 0 || double.IsNaN(options.Width.Value) || double.IsInfinity(options.Width.Value))) {
            throw new ArgumentOutOfRangeException(nameof(options), "Image stamp width must be a positive finite value.");
        }

        if (options.Height.HasValue && (options.Height.Value <= 0 || double.IsNaN(options.Height.Value) || double.IsInfinity(options.Height.Value))) {
            throw new ArgumentOutOfRangeException(nameof(options), "Image stamp height must be a positive finite value.");
        }

        if ((options.X.HasValue && (double.IsNaN(options.X.Value) || double.IsInfinity(options.X.Value))) ||
            (options.Y.HasValue && (double.IsNaN(options.Y.Value) || double.IsInfinity(options.Y.Value))) ||
            double.IsNaN(options.RotationDegrees) ||
            double.IsInfinity(options.RotationDegrees)) {
            throw new ArgumentOutOfRangeException(nameof(options), "Image stamp coordinates and rotation must be finite.");
        }
    }
}
