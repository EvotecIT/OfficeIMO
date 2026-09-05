namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private static string ResolveRubyPosition(string? value, string? inherited) {
        string normalized = string.IsNullOrWhiteSpace(value)
            ? string.IsNullOrWhiteSpace(inherited) ? "over" : inherited!
            : value!.Trim().ToLowerInvariant();
        return normalized == "under" ? "under" : "over";
    }

    private static string ResolveRubyAlign(string? value, string? inherited) {
        string normalized = string.IsNullOrWhiteSpace(value)
            ? string.IsNullOrWhiteSpace(inherited) ? "space-around" : inherited!
            : value!.Trim().ToLowerInvariant();
        return normalized == "start" || normalized == "center" || normalized == "space-between"
            ? normalized
            : "space-around";
    }
}
