namespace OfficeIMO.Core.Internal;

/// <summary>Accessibility text rules shared by the Word, PowerPoint and Excel checkers.</summary>
internal static class OfficeAccessibilityText {
    private static readonly string[] VagueLinkLabels = { "click here", "here", "more", "link", "read more" };

    /// <summary>Whether hyperlink text describes its destination without surrounding context.</summary>
    internal static bool IsMeaningfulLinkLabel(string? value) {
        if (string.IsNullOrWhiteSpace(value)) return false;
        string normalized = value!.Trim();
        if (normalized.Length < 3) return false;
        if (Uri.TryCreate(normalized, UriKind.Absolute, out _)) return false;
        foreach (string vague in VagueLinkLabels) {
            if (string.Equals(normalized, vague, StringComparison.OrdinalIgnoreCase)) return false;
        }
        return true;
    }
}
