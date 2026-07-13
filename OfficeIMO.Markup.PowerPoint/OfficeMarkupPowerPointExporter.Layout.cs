using System.Diagnostics;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Markup.PowerPoint;

internal sealed partial class OfficeMarkupPowerPointExporter {
    private static LayoutCursor ResolveBox(
        OfficeMarkupPlacement? placement,
        IDictionary<string, string> attributes,
        LayoutCursor fallback,
        double defaultHeight,
        SlideCanvasMetrics metrics) {
        if (!HasExplicitPlacement(placement, attributes)) {
            return new LayoutCursor(fallback.Left, fallback.Top, fallback.Width, Math.Min(metrics.Vertical(defaultHeight), fallback.RemainingHeight));
        }

        var left = ParsePercentOrInches(PlacementValue(placement, attributes, "x"), fallback.Left, metrics.Width);
        var top = ParsePercentOrInches(PlacementValue(placement, attributes, "y"), fallback.Top, metrics.Height);
        var width = ParsePercentOrInches(PlacementValue(placement, attributes, "w"), fallback.Width, metrics.Width);
        var height = ParsePercentOrInches(PlacementValue(placement, attributes, "h"), Math.Min(metrics.Vertical(defaultHeight), fallback.RemainingHeight), metrics.Height);
        return new LayoutCursor(left, top, width, height);
    }

    private static LayoutCursor ResolveBox(IDictionary<string, string> attributes, LayoutCursor fallback, double defaultHeight, SlideCanvasMetrics metrics) =>
        ResolveBox(null, attributes, fallback, defaultHeight, metrics);

    private static bool HasExplicitPlacement(OfficeMarkupBlock block) =>
        HasExplicitPlacement(GetPlacement(block), block.Attributes);

    private static bool HasExplicitPlacement(OfficeMarkupPlacement? placement, IDictionary<string, string> attributes) =>
        placement?.HasValue == true || HasExplicitPlacement(attributes);

    private static bool HasExplicitPlacement(IDictionary<string, string> attributes) =>
        attributes.ContainsKey("x")
        || attributes.ContainsKey("y")
        || attributes.ContainsKey("w")
        || attributes.ContainsKey("h")
        || attributes.ContainsKey("width")
        || attributes.ContainsKey("height");

    private static string? PlacementValue(OfficeMarkupPlacement? placement, IDictionary<string, string> attributes, string name) {
        if (placement != null) {
            switch (name) {
                case "x":
                    if (!string.IsNullOrWhiteSpace(placement.X)) {
                        return placement.X;
                    }
                    break;
                case "y":
                    if (!string.IsNullOrWhiteSpace(placement.Y)) {
                        return placement.Y;
                    }
                    break;
                case "w":
                    if (!string.IsNullOrWhiteSpace(placement.Width)) {
                        return placement.Width;
                    }
                    break;
                case "h":
                    if (!string.IsNullOrWhiteSpace(placement.Height)) {
                        return placement.Height;
                    }
                    break;
            }
        }

        if (attributes.TryGetValue(name, out var value)) {
            return value;
        }

        if (name == "w" && attributes.TryGetValue("width", out value)) {
            return value;
        }

        if (name == "h" && attributes.TryGetValue("height", out value)) {
            return value;
        }

        return null;
    }

    private static double ParsePercentOrInches(string? value, double fallback, double size) {
        if (string.IsNullOrWhiteSpace(value)) {
            return fallback;
        }

        value = value!.Trim();
        if (value.EndsWith("%", StringComparison.Ordinal)) {
            return double.TryParse(value.Substring(0, value.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out var percent)
                ? size * (percent / 100.0)
                : fallback;
        }

        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var inches) ? inches : fallback;
    }

    private static void ApplyTransition(PowerPointSlide slide, string? transition) {
        if (string.IsNullOrWhiteSpace(transition)) {
            return;
        }

        var resolvedTransition = OfficeMarkupTransitionResolver.Parse(transition);
        if (string.IsNullOrWhiteSpace(resolvedTransition.ResolvedIdentifier)) {
            return;
        }

        if (Enum.TryParse<SlideTransition>(resolvedTransition.ResolvedIdentifier, true, out var parsed)) {
            slide.Transition = parsed;
            ApplyTransitionAttributes(slide, resolvedTransition.Attributes);
        }
    }

    private static void ApplyTransitionAttributes(PowerPointSlide slide, IReadOnlyDictionary<string, string> attributes) {
        if (TryGetTransitionSpeed(attributes, out var speed)) {
            slide.TransitionSpeed = speed;
        }

        if (TryGetTransitionSeconds(attributes, out var durationSeconds, "duration", "dur")) {
            slide.TransitionDurationSeconds = durationSeconds;
        }

        if (TryGetTransitionBoolean(attributes, out var advanceOnClick, "advance-on-click", "advanceonclick", "advance-click", "onclick", "click")) {
            slide.TransitionAdvanceOnClick = advanceOnClick;
        }

        if (TryGetTransitionSeconds(attributes, out var advanceAfterSeconds, "advance-after", "advanceafter", "after", "delay")) {
            slide.TransitionAdvanceAfterSeconds = advanceAfterSeconds;
        }
    }

    private static bool TryGetTransitionSpeed(IReadOnlyDictionary<string, string> attributes, out SlideTransitionSpeed speed) {
        speed = default;
        var value = GetTransitionAttribute(attributes, "speed", "spd");
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        switch (NormalizeTransitionToken(value)) {
            case "slow":
                speed = SlideTransitionSpeed.Slow;
                return true;
            case "medium":
            case "med":
                speed = SlideTransitionSpeed.Medium;
                return true;
            case "fast":
                speed = SlideTransitionSpeed.Fast;
                return true;
            default:
                return false;
        }
    }

    private static bool TryGetTransitionSeconds(IReadOnlyDictionary<string, string> attributes, out double seconds, params string[] names) {
        seconds = default;
        var value = GetTransitionAttribute(attributes, names);
        return !string.IsNullOrWhiteSpace(value)
               && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out seconds);
    }

    private static bool TryGetTransitionBoolean(IReadOnlyDictionary<string, string> attributes, out bool enabled, params string[] names) {
        enabled = default;
        var value = GetTransitionAttribute(attributes, names);
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        switch (NormalizeTransitionToken(value)) {
            case "true":
            case "yes":
            case "on":
            case "1":
                enabled = true;
                return true;
            case "false":
            case "no":
            case "off":
            case "0":
                enabled = false;
                return true;
            default:
                return false;
        }
    }

    private static string? GetTransitionAttribute(IReadOnlyDictionary<string, string> attributes, params string[] names) {
        foreach (var name in names) {
            if (attributes.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)) {
                return value.Trim();
            }
        }

        return null;
    }

    private static string NormalizeTransitionToken(string? value) =>
        new string((value ?? string.Empty).Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());

    private static IReadOnlyList<OfficeMarkupBlock> ParseLightweightMarkdown(string body) {
        var blocks = new List<OfficeMarkupBlock>();
        foreach (var rawLine in body.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')) {
            var line = rawLine.Trim();
            if (line.Length == 0) {
                continue;
            }

            if (line.StartsWith("### ", StringComparison.Ordinal)) {
                blocks.Add(new OfficeMarkupHeadingBlock(3, line.Substring(4)));
            } else if (line.StartsWith("## ", StringComparison.Ordinal)) {
                blocks.Add(new OfficeMarkupHeadingBlock(2, line.Substring(3)));
            } else if (line.StartsWith("# ", StringComparison.Ordinal)) {
                blocks.Add(new OfficeMarkupHeadingBlock(1, line.Substring(2)));
            } else if (line.StartsWith("- ", StringComparison.Ordinal)) {
                var list = blocks.LastOrDefault() as OfficeMarkupListBlock;
                if (list == null) {
                    list = new OfficeMarkupListBlock(false);
                    blocks.Add(list);
                }

                list.Items.Add(new OfficeMarkupListItem(line.Substring(2)));
            } else {
                blocks.Add(new OfficeMarkupParagraphBlock(line));
            }
        }

        return blocks;
    }

    private static bool IsExtension(OfficeMarkupBlock block, string command) =>
        block is OfficeMarkupExtensionBlock extension
        && string.Equals(Normalize(extension.Command), Normalize(command), StringComparison.Ordinal);

    private static bool IsColumns(OfficeMarkupBlock block) =>
        block is OfficeMarkupColumnsBlock || IsExtension(block, "columns");

    private static bool IsColumn(OfficeMarkupBlock block) =>
        block is OfficeMarkupColumnBlock
        || IsExtension(block, "column")
        || IsExtension(block, "left")
        || IsExtension(block, "right");

    private static string GetColumnBody(OfficeMarkupBlock block) {
        if (block is OfficeMarkupColumnBlock column) {
            return column.Body;
        }

        return block is OfficeMarkupExtensionBlock extension ? extension.Body : string.Empty;
    }

    private static OfficeMarkupPlacement? GetPlacement(OfficeMarkupBlock block) {
        switch (block) {
            case OfficeMarkupImageBlock image:
                return image.Placement;
            case OfficeMarkupDiagramBlock diagram:
                return diagram.Placement;
            case OfficeMarkupChartBlock chart:
                return chart.Placement;
            case OfficeMarkupTextBoxBlock textBox:
                return textBox.Placement;
            case OfficeMarkupColumnsBlock columns:
                return columns.Placement;
            case OfficeMarkupCardBlock card:
                return card.Placement;
            default:
                return null;
        }
    }

    private static double ResolveGap(OfficeMarkupBlock block, SlideCanvasMetrics metrics) {
        string? value = null;
        if (block is OfficeMarkupColumnsBlock columns) {
            value = columns.Gap;
        }

        if (string.IsNullOrWhiteSpace(value)) {
            value = GetAttribute(block.Attributes, "gap");
        }

        return ParsePercentOrInches(value, metrics.Horizontal(0.28), metrics.Width);
    }

    private static string? GetAttribute(IDictionary<string, string> attributes, string name) =>
        attributes.TryGetValue(name, out var value) ? value : null;

    private static string? GetAttribute(IDictionary<string, string> attributes, params string[] names) {
        foreach (var name in names) {
            if (attributes.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)) {
                return value.Trim();
            }
        }

        return null;
    }

    private static string? NormalizeSectionName(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        return value!.Trim();
    }

    private static string Normalize(string value) => (value ?? string.Empty).Replace("-", string.Empty).ToLowerInvariant();

    private static double EstimateTextHeight(string text) {
        var lines = Math.Max(1, text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n').Length);
        return Math.Min(1.4, 0.3 * lines + 0.2);
    }

    private static void TryDelete(string path) {
        try {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        } catch (IOException ex) {
            Trace.TraceWarning($"OfficeIMO.Markup.PowerPoint could not delete temporary file '{path}': {ex.Message}");
        } catch (UnauthorizedAccessException ex) {
            Trace.TraceWarning($"OfficeIMO.Markup.PowerPoint could not delete temporary file '{path}': {ex.Message}");
        }
    }

    private sealed class LayoutCursor {
        public LayoutCursor(double left, double top, double width, double height) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
            InitialTop = top;
        }

        public double Left { get; }
        public double Top { get; private set; }
        public double Width { get; }
        public double Height { get; }
        public double Bottom => Top + Height;
        public double RemainingHeight => Math.Max(0.28, (InitialTop + Height) - Top);
        private double InitialTop { get; }

        public void Advance(double height) {
            Top += height + 0.12;
        }

        public void MoveToBottom() {
            Top = InitialTop + Height;
        }
    }
}
