using System.Globalization;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>Validates the one-based page numbers entered in desktop and touch readers.</summary>
internal static class PdfPageNavigation {
    internal static bool TryParsePageNumber(string? text, int pageCount, out int pageNumber) =>
        int.TryParse(text, NumberStyles.Integer, CultureInfo.CurrentCulture, out pageNumber) &&
        pageNumber >= 1 && pageNumber <= pageCount;
}
