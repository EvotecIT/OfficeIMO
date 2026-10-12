using System.Text;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Builds file-picker glob patterns that match extensions regardless of letter case.</summary>
internal static class StudioPickerPatterns {
    // GTK and portal pickers match globs case-sensitively; SCAN.PDF, IMG_0001.JPG or report.Pdf would otherwise be hidden.
    internal static string[] CaseInsensitive(params string[] patterns) =>
        CaseInsensitive(patterns, OperatingSystem.IsWindows() || OperatingSystem.IsMacOS() || OperatingSystem.IsIOS());

    internal static string[] CaseInsensitive(IEnumerable<string> patterns, bool platformIgnoresCase) =>
        platformIgnoresCase ? patterns.ToArray() : patterns.Select(AnyCase).Distinct(StringComparer.Ordinal).ToArray();

    // "*.pdf" becomes "*.[pP][dD][fF]", which matches every mix of upper and lower case.
    private static string AnyCase(string pattern) {
        var glob = new StringBuilder(pattern.Length * 4);
        foreach (char character in pattern) {
            char lower = char.ToLowerInvariant(character), upper = char.ToUpperInvariant(character);
            if (lower == upper) glob.Append(character);
            else glob.Append('[').Append(lower).Append(upper).Append(']');
        }
        return glob.ToString();
    }
}
