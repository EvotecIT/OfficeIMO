using System.Globalization;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>Resolves printer media keywords to physical sizes and back.</summary>
internal static class PdfPrinterMediaNames {
    // PPD keywords, in points. "B4"/"B5" are JIS sizes in PPD files; ISO B uses "ISOB4"/"ISOB5".
    private static readonly Dictionary<string, (double Width, double Height)> PpdSizes = new(StringComparer.OrdinalIgnoreCase) {
        ["Letter"] = (612, 792), ["Legal"] = (612, 1008), ["Executive"] = (522, 756), ["Tabloid"] = (792, 1224),
        ["Ledger"] = (792, 1224), ["Statement"] = (396, 612), ["FanFoldGermanLegal"] = (612, 936),
        ["A3"] = (841.89, 1190.55), ["A4"] = (595.28, 841.89), ["A5"] = (419.53, 595.28), ["A6"] = (297.64, 419.53),
        ["B4"] = (728.5, 1031.81), ["B5"] = (515.91, 728.5), ["B6"] = (362.83, 515.91),
        ["ISOB4"] = (708.66, 1000.63), ["ISOB5"] = (498.9, 708.66),
        ["Env10"] = (296.79, 684), ["EnvDL"] = (311.81, 623.62), ["EnvC5"] = (459.21, 649.13), ["EnvC4"] = (649.13, 918.43),
        ["EnvMonarch"] = (279, 540), ["Postcard"] = (283.46, 419.53)
    };

    private static readonly Regex Dimensions = new(@"^(?:[a-z0-9-]+_[a-z0-9.-]+_)?(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)(mm|in)?$",
        RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, TimeSpan.FromSeconds(1));

    /// <summary>Returns the portrait size of a PPD keyword or PWG self-describing media name, or null when unknown.</summary>
    internal static PdfPageSize? TryGetSize(string keyword) {
        if (keyword.StartsWith("Custom.", StringComparison.OrdinalIgnoreCase)) return null;
        Match match = Dimensions.Match(keyword);
        if (match.Success) {
            double width = double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
            double height = double.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
            string unit = match.Groups[3].Value.ToLowerInvariant();
            // Unitless PPD names such as 4x6 are inches.
            double scale = unit == "mm" ? 72D / 25.4D : 72D;
            if (unit.Length == 0 && (width > 60 || height > 60)) return null;
            return Portrait(width * scale, height * scale);
        }
        // PPD variants such as A4.Borderless keep the base size.
        int variant = keyword.IndexOf('.');
        string name = variant > 0 ? keyword.Substring(0, variant) : keyword;
        return PpdSizes.TryGetValue(name, out var size) ? Portrait(size.Width, size.Height) : null;
    }

    /// <summary>Finds the media whose size matches the paper within two points, preferring the default and plain keywords.</summary>
    internal static PdfPrinterMedia? Match(IReadOnlyList<PdfPrinterMedia> media, PdfPrinterMedia? preferred, PdfPageSize paper) {
        if (preferred?.Size is { } preferredSize && SameSize(preferredSize, paper)) return preferred;
        PdfPrinterMedia? variant = null;
        foreach (PdfPrinterMedia item in media) {
            if (item.Size is not { } size || !SameSize(size, paper)) continue;
            if (item.Id.IndexOf('.') < 0) return item;
            variant ??= item;
        }
        return variant;
    }

    /// <summary>Returns the standard PPD keyword for a paper size, or null when it has none.</summary>
    internal static string? TryGetKeyword(PdfPageSize paper) {
        foreach (KeyValuePair<string, (double Width, double Height)> entry in PpdSizes)
            if (entry.Key != "Ledger" && SameSize(paper, Portrait(entry.Value.Width, entry.Value.Height))) return entry.Key;
        return null;
    }

    internal static bool SameSize(PdfPageSize first, PdfPageSize second) {
        double firstShort = Math.Min(first.Width.ToPoints(), first.Height.ToPoints()), firstLong = Math.Max(first.Width.ToPoints(), first.Height.ToPoints());
        double secondShort = Math.Min(second.Width.ToPoints(), second.Height.ToPoints()), secondLong = Math.Max(second.Width.ToPoints(), second.Height.ToPoints());
        return Math.Abs(firstShort - secondShort) <= 2 && Math.Abs(firstLong - secondLong) <= 2;
    }

    internal static PdfPageSize Portrait(double width, double height) =>
        new(OfficeLength.Points(Math.Min(width, height)), OfficeLength.Points(Math.Max(width, height)));
}
