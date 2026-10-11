namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Builds file-picker glob patterns that match extensions regardless of letter case.</summary>
internal static class StudioPickerPatterns {
    // GTK and portal pickers match globs case-sensitively; SCAN.PDF or IMG_0001.JPG would otherwise be hidden.
    internal static string[] CaseInsensitive(params string[] patterns) =>
        CaseInsensitive(patterns, OperatingSystem.IsWindows() || OperatingSystem.IsMacOS() || OperatingSystem.IsIOS());

    internal static string[] CaseInsensitive(IEnumerable<string> patterns, bool platformIgnoresCase) {
        if (platformIgnoresCase) return patterns.ToArray();
        return patterns
            .SelectMany(static pattern => new[] { pattern, pattern.ToLowerInvariant(), pattern.ToUpperInvariant() })
            .Distinct(StringComparer.Ordinal)
            .ToArray();
    }
}
