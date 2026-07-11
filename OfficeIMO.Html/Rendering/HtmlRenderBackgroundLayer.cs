using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlRenderBackgroundLayer {
    internal HtmlRenderBackgroundLayer(string source, string position, string repeat, string size) {
        Source = source;
        Position = string.IsNullOrWhiteSpace(position) ? "0% 0%" : position;
        Repeat = string.IsNullOrWhiteSpace(repeat) ? "repeat" : repeat;
        Size = string.IsNullOrWhiteSpace(size) ? "auto" : size;
    }

    internal HtmlRenderBackgroundLayer(HtmlCssLinearGradientDefinition gradient, string position, string repeat, string size) {
        LinearGradient = gradient ?? throw new ArgumentNullException(nameof(gradient));
        Position = string.IsNullOrWhiteSpace(position) ? "0% 0%" : position;
        Repeat = string.IsNullOrWhiteSpace(repeat) ? "repeat" : repeat;
        Size = string.IsNullOrWhiteSpace(size) ? "auto" : size;
    }

    internal HtmlRenderBackgroundLayer(HtmlCssRadialGradientDefinition gradient, string position, string repeat, string size) {
        RadialGradient = gradient ?? throw new ArgumentNullException(nameof(gradient));
        Position = string.IsNullOrWhiteSpace(position) ? "0% 0%" : position;
        Repeat = string.IsNullOrWhiteSpace(repeat) ? "repeat" : repeat;
        Size = string.IsNullOrWhiteSpace(size) ? "auto" : size;
    }

    internal string? Source { get; }
    internal HtmlCssLinearGradientDefinition? LinearGradient { get; }
    internal HtmlCssRadialGradientDefinition? RadialGradient { get; }
    internal string Position { get; }
    internal string Repeat { get; }
    internal string Size { get; }
}
