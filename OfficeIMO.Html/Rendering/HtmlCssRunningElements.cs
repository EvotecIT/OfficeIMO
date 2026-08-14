using System.Globalization;
using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal static class HtmlCssRunningElementKeys {
    private const string Prefix = "\0officeimo-running-element:";

    internal static string ForName(string name) => Prefix + name;
}

internal sealed class HtmlCssRunningElementSnapshot {
    internal HtmlCssRunningElementSnapshot(HtmlRenderFlowBlock block, IElement element, HtmlRenderBoxStyle parentStyle, int depth) {
        Block = block ?? throw new ArgumentNullException(nameof(block));
        Element = element ?? throw new ArgumentNullException(nameof(element));
        ParentStyle = parentStyle?.Clone() ?? throw new ArgumentNullException(nameof(parentStyle));
        Depth = depth;
    }

    internal HtmlRenderFlowBlock Block { get; }
    internal IElement Element { get; }
    internal HtmlRenderBoxStyle ParentStyle { get; }
    internal int Depth { get; }
}

internal static class HtmlCssRunningElementParser {
    internal static bool TryParsePosition(string? value, out string name) {
        name = string.Empty;
        string normalized = value?.Trim() ?? string.Empty;
        const string prefix = "running(";
        if (normalized.Length <= prefix.Length
            || !normalized.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return false;
        int open = prefix.Length - 1;
        int close = HtmlRenderCssValues.FindMatchingParenthesis(normalized, open);
        if (close != normalized.Length - 1) return false;
        return HtmlCssIdentifierParser.TryParse(normalized.Substring(open + 1, close - open - 1).Trim(), out name);
    }

    internal static string FormatSnapshotId(int id) => id.ToString(CultureInfo.InvariantCulture);

    internal static bool TryParseSnapshotId(string? value, out int id) =>
        int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out id) && id > 0;
}
