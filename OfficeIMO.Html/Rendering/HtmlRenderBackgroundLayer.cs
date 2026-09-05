using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlRenderBackgroundLayer {
    internal HtmlRenderBackgroundLayer(string source, string position, string repeat, string size, string origin, string clip, string attachment) {
        Source = source;
        Initialize(position, repeat, size, origin, clip, attachment);
    }

    internal HtmlRenderBackgroundLayer(HtmlCssLinearGradientDefinition gradient, string position, string repeat, string size, string origin, string clip, string attachment) {
        LinearGradient = gradient ?? throw new ArgumentNullException(nameof(gradient));
        Initialize(position, repeat, size, origin, clip, attachment);
    }

    internal HtmlRenderBackgroundLayer(HtmlCssRadialGradientDefinition gradient, string position, string repeat, string size, string origin, string clip, string attachment) {
        RadialGradient = gradient ?? throw new ArgumentNullException(nameof(gradient));
        Initialize(position, repeat, size, origin, clip, attachment);
    }

    internal HtmlRenderBackgroundLayer(HtmlCssConicGradientDefinition gradient, string position, string repeat, string size, string origin, string clip, string attachment) {
        ConicGradient = gradient ?? throw new ArgumentNullException(nameof(gradient));
        Initialize(position, repeat, size, origin, clip, attachment);
    }

    private void Initialize(string position, string repeat, string size, string origin, string clip, string attachment) {
        Position = string.IsNullOrWhiteSpace(position) ? "0% 0%" : position;
        Repeat = string.IsNullOrWhiteSpace(repeat) ? "repeat" : repeat;
        Size = string.IsNullOrWhiteSpace(size) ? "auto" : size;
        Origin = NormalizeBox(origin, "padding-box");
        Clip = NormalizeBox(clip, "border-box");
        Attachment = NormalizeAttachment(attachment);
    }

    internal string? Source { get; }
    internal HtmlCssLinearGradientDefinition? LinearGradient { get; }
    internal HtmlCssRadialGradientDefinition? RadialGradient { get; }
    internal HtmlCssConicGradientDefinition? ConicGradient { get; }
    internal string Position { get; private set; } = "0% 0%";
    internal string Repeat { get; private set; } = "repeat";
    internal string Size { get; private set; } = "auto";
    internal string Origin { get; private set; } = "padding-box";
    internal string Clip { get; private set; } = "border-box";
    internal string Attachment { get; private set; } = "scroll";

    internal static string NormalizeBox(string value, string fallback) {
        string normalized = value.Trim().ToLowerInvariant();
        return normalized == "border-box" || normalized == "padding-box" || normalized == "content-box"
            ? normalized
            : fallback;
    }

    private static string NormalizeAttachment(string value) {
        string normalized = value.Trim().ToLowerInvariant();
        return normalized == "fixed" || normalized == "local" || normalized == "scroll" ? normalized : "scroll";
    }
}
