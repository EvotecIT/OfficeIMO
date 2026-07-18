using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageEditor {
    private static void ValidatePageNumbers(int[] pageNumbers, int pageCount, string paramName, bool allowDuplicates = false) {
        var seen = new HashSet<int>();
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            if (pageNumber < 1 || pageNumber > pageCount) {
                throw new ArgumentOutOfRangeException(paramName, "Page number " + pageNumber.ToString(CultureInfo.InvariantCulture) + " is outside the document page range 1-" + pageCount.ToString(CultureInfo.InvariantCulture) + ".");
            }

            if (!allowDuplicates && !seen.Add(pageNumber)) {
                throw new ArgumentException("Duplicate page selections are not supported.", paramName);
            }
        }
    }

    private static void ValidateReorderPageNumbers(int[] pageNumbers, int pageCount, string paramName) {
        if (pageNumbers.Length == 0) {
            throw new ArgumentException("At least one page number must be specified.", paramName);
        }

        if (pageNumbers.Length != pageCount) {
            throw new ArgumentException("Reorder must include every page exactly once.", paramName);
        }

        ValidatePageNumbers(pageNumbers, pageCount, paramName);
    }

    private static int[] ExpandPageRanges(PdfPageRange[] pageRanges, string paramName) {
        Guard.NotNull(pageRanges, paramName);
        if (pageRanges.Length == 0) {
            throw new ArgumentException("At least one page range must be specified.", paramName);
        }

        var pages = new List<int>();
        for (int i = 0; i < pageRanges.Length; i++) {
            pages.AddRange(pageRanges[i].ToPageNumbers());
        }

        return pages.ToArray();
    }

    private static int[] ExpandPageRangesDistinct(PdfPageRange[] pageRanges, string paramName) {
        int[] pages = ExpandPageRanges(pageRanges, paramName);
        var seen = new HashSet<int>();
        var distinct = new List<int>(pages.Length);
        for (int i = 0; i < pages.Length; i++) {
            if (seen.Add(pages[i])) {
                distinct.Add(pages[i]);
            }
        }

        return distinct.ToArray();
    }

    private static void ValidateMoveInsertBeforePageNumber(int insertBeforePageNumber, int pageCount) {
        if (insertBeforePageNumber < 1 || insertBeforePageNumber > pageCount + 1) {
            throw new ArgumentOutOfRangeException(nameof(insertBeforePageNumber), "Insert-before page must be in the document page range 1-" + (pageCount + 1).ToString(CultureInfo.InvariantCulture) + ".");
        }
    }

    private static int[] BuildInclusivePageRange(int firstPage, int lastPage, string lastPageParamName) {
        if (firstPage > lastPage) {
            throw new ArgumentOutOfRangeException(lastPageParamName, "Last page must be greater than or equal to first page.");
        }

        return Enumerable.Range(firstPage, lastPage - firstPage + 1).ToArray();
    }

    private static int NormalizeRotation(int rotationDegrees) {
        if (rotationDegrees % 90 != 0) {
            throw new ArgumentOutOfRangeException(nameof(rotationDegrees), "Rotation must be a multiple of 90 degrees.");
        }

        int normalized = rotationDegrees % 360;
        if (normalized < 0) {
            normalized += 360;
        }

        return normalized;
    }
}
