using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer.IntelligenceX;

/// <summary>
/// Typed IntelligenceX visual fence metadata layered on top of <see cref="MarkdownCodeFenceInfo"/>.
/// </summary>
public sealed class IntelligenceXVisualFenceOptions {
    private static readonly IReadOnlyList<string> EmptyClasses = Array.Empty<string>();

    private IntelligenceXVisualFenceOptions(
        string language,
        string infoString,
        string? elementId,
        IReadOnlyList<string> classes,
        string? title,
        bool pinned,
        string? theme,
        string? variant,
        string? view,
        int? maxItems) {
        Language = language ?? string.Empty;
        InfoString = infoString ?? string.Empty;
        ElementId = elementId;
        Classes = classes ?? EmptyClasses;
        Title = title;
        Pinned = pinned;
        Theme = NormalizeOptional(theme);
        Variant = NormalizeOptional(variant);
        View = NormalizeOptional(view);
        MaxItems = maxItems;
    }

    /// <summary>Primary fence language token.</summary>
    public string Language { get; }

    /// <summary>Full normalized fence info string.</summary>
    public string InfoString { get; }

    /// <summary>Optional element id parsed from the source fence metadata.</summary>
    public string? ElementId { get; }

    /// <summary>Optional CSS classes parsed from the source fence metadata.</summary>
    public IReadOnlyList<string> Classes { get; }

    /// <summary>Optional visual title resolved from the source fence metadata.</summary>
    public string? Title { get; }

    /// <summary>Whether the fence metadata explicitly pins the visual in host layouts.</summary>
    public bool Pinned { get; }

    /// <summary>Optional theme or palette selector.</summary>
    public string? Theme { get; }

    /// <summary>Optional named variant selector.</summary>
    public string? Variant { get; }

    /// <summary>Optional named view or mode selector.</summary>
    public string? View { get; }

    /// <summary>Optional host-defined item limit parsed from common aliases.</summary>
    public int? MaxItems { get; }

    /// <summary>
    /// Determines whether the parsed metadata contains the given CSS class.
    /// </summary>
    public bool HasClass(string className) {
        if (string.IsNullOrWhiteSpace(className) || Classes.Count == 0) {
            return false;
        }

        for (int i = 0; i < Classes.Count; i++) {
            if (string.Equals(Classes[i], className.Trim(), StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Parses typed IntelligenceX visual fence metadata from a raw fenced-code info string.
    /// </summary>
    public static IntelligenceXVisualFenceOptions Parse(string? infoString) {
        return Parse(MarkdownCodeFenceInfo.Parse(infoString));
    }

    /// <summary>
    /// Parses typed IntelligenceX visual fence metadata from a shared fenced-code info descriptor.
    /// </summary>
    public static IntelligenceXVisualFenceOptions Parse(MarkdownCodeFenceInfo? fenceInfo) {
        if (fenceInfo == null) {
            return new IntelligenceXVisualFenceOptions(
                string.Empty,
                string.Empty,
                null,
                EmptyClasses,
                null,
                pinned: false,
                theme: null,
                variant: null,
                view: null,
                maxItems: null);
        }

        var parsedOptions = IntelligenceXVisualFenceSchemas.Visuals.Parse(fenceInfo);
        return new IntelligenceXVisualFenceOptions(
            fenceInfo.Language,
            fenceInfo.InfoString,
            fenceInfo.ElementId,
            fenceInfo.Classes,
            fenceInfo.Title,
            parsedOptions.TryGetBoolean("pinned", out var pinned) && pinned,
            parsedOptions.TryGetString("theme", out var theme) ? theme : null,
            parsedOptions.TryGetString("variant", out var variant) ? variant : null,
            parsedOptions.TryGetString("view", out var view) ? view : null,
            parsedOptions.TryGetInt32("maxItems", out var maxItems)
                ? maxItems
                : (int?) null);
    }

    private static string? NormalizeOptional(string? value) {
        return string.IsNullOrWhiteSpace(value) ? null : value!.Trim();
    }
}
