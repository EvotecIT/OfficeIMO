using System.Globalization;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    internal static bool IsSupportedPdfTagType(string? value) {
        string normalized = value?.Trim().ToLowerInvariant() ?? string.Empty;
        return normalized.Length == 0
            || normalized == "auto" || normalized == "artifact" || normalized == "none"
            || TryParsePdfSemanticRole(normalized, out _);
    }

    private static void ApplyPdfSemanticTag(string? value, HtmlRenderBoxStyle style) {
        string normalized = value?.Trim().ToLowerInvariant() ?? string.Empty;
        if (normalized.Length == 0 || normalized == "auto") return;
        if (normalized == "artifact" || normalized == "none") {
            style.SemanticArtifact = true;
            return;
        }
        if (TryParsePdfSemanticRole(normalized, out HtmlRenderSemanticGroupRole role)) {
            style.SemanticGroupRoleOverride = role;
            return;
        }
        style.UnsupportedSemanticTag = normalized;
    }

    private static bool TryParsePdfSemanticRole(string value, out HtmlRenderSemanticGroupRole role) {
        switch (value) {
            case "sect": case "section": role = HtmlRenderSemanticGroupRole.Section; return true;
            case "div": case "division": role = HtmlRenderSemanticGroupRole.Division; return true;
            case "p": case "paragraph": role = HtmlRenderSemanticGroupRole.Paragraph; return true;
            case "h1": role = HtmlRenderSemanticGroupRole.Heading1; return true;
            case "h2": role = HtmlRenderSemanticGroupRole.Heading2; return true;
            case "h3": role = HtmlRenderSemanticGroupRole.Heading3; return true;
            case "h4": role = HtmlRenderSemanticGroupRole.Heading4; return true;
            case "h5": role = HtmlRenderSemanticGroupRole.Heading5; return true;
            case "h6": role = HtmlRenderSemanticGroupRole.Heading6; return true;
            case "l": case "list": role = HtmlRenderSemanticGroupRole.List; return true;
            case "li": case "list-item": role = HtmlRenderSemanticGroupRole.ListItem; return true;
            case "lbl": case "list-label": role = HtmlRenderSemanticGroupRole.ListLabel; return true;
            case "lbody": case "list-body": role = HtmlRenderSemanticGroupRole.ListBody; return true;
            case "table": role = HtmlRenderSemanticGroupRole.Table; return true;
            case "tr": case "table-row": role = HtmlRenderSemanticGroupRole.TableRow; return true;
            case "th": case "table-header-cell": role = HtmlRenderSemanticGroupRole.TableHeaderCell; return true;
            case "td": case "table-cell": role = HtmlRenderSemanticGroupRole.TableCell; return true;
            case "caption": role = HtmlRenderSemanticGroupRole.Caption; return true;
            default: role = default; return false;
        }
    }

    private static void ApplyBookmark(HtmlComputedStyle computed, HtmlRenderBoxStyle style) {
        string level = computed.GetValue("bookmark-level").Trim().ToLowerInvariant();
        string label = computed.GetValue("bookmark-label").Trim();
        string state = computed.GetValue("bookmark-state").Trim().ToLowerInvariant();
        if (level.Length > 0) {
            style.BookmarkLevelSpecified = true;
            if (level == "none") style.BookmarkSuppressed = true;
            else if (int.TryParse(level, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 1 && parsed <= 64) style.BookmarkLevel = parsed;
            else style.UnsupportedBookmark = "bookmark-level=" + level;
        }
        if (label.Length > 0) {
            if (string.Equals(label, "content(text)", StringComparison.OrdinalIgnoreCase)) style.BookmarkLabel = null;
            else if (label.Length >= 2 && (label[0] == '\'' || label[0] == '"') && label[label.Length - 1] == label[0]) {
                style.BookmarkLabel = HtmlCssEscapeDecoder.Decode(label.Substring(1, label.Length - 2));
            } else style.UnsupportedBookmark = "bookmark-label=" + label;
        }
        if (state.Length > 0) {
            if (state == "open") style.BookmarkState = HtmlRenderBookmarkState.Open;
            else if (state == "closed") style.BookmarkState = HtmlRenderBookmarkState.Closed;
            else style.UnsupportedBookmark = "bookmark-state=" + state;
        }
    }
}
