using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed class HtmlRenderBackgroundLayer {
    internal HtmlRenderBackgroundLayer(string source, string position, string repeat, string size) {
        Source = source;
        Position = string.IsNullOrWhiteSpace(position) ? "0% 0%" : position;
        Repeat = string.IsNullOrWhiteSpace(repeat) ? "repeat" : repeat;
        Size = string.IsNullOrWhiteSpace(size) ? "auto" : size;
    }

    internal HtmlRenderBackgroundLayer(OfficeLinearGradient gradient, string position, string repeat, string size) {
        LinearGradient = gradient?.Clone() ?? throw new ArgumentNullException(nameof(gradient));
        Position = string.IsNullOrWhiteSpace(position) ? "0% 0%" : position;
        Repeat = string.IsNullOrWhiteSpace(repeat) ? "repeat" : repeat;
        Size = string.IsNullOrWhiteSpace(size) ? "auto" : size;
    }

    internal string? Source { get; }
    internal OfficeLinearGradient? LinearGradient { get; }
    internal string Position { get; }
    internal string Repeat { get; }
    internal string Size { get; }
}
