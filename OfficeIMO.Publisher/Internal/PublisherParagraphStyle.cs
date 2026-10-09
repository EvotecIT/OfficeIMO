using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

// Nullable native values distinguish an explicit override from an inherited value.
internal sealed class PublisherParagraphStyle {
    internal uint? DefaultStyleIndex { get; set; }
    internal OfficeTextAlignment? Alignment { get; set; }
    internal double? LineHeight { get; set; }
    internal double? LineHeightFactor { get; set; }
    internal double? Before { get; set; }
    internal double? After { get; set; }
    internal double? Left { get; set; }
    internal double? Right { get; set; }
    internal double? FirstLine { get; set; }
    internal IReadOnlyList<OfficeTextTabStop>? Tabs { get; set; }
    internal PublisherListStyle? List { get; set; }
    internal double? LabelSize { get; set; }
    internal uint? LabelFont { get; set; }
    internal double? LabelTextPosition { get; set; }

    internal OfficeTextPadding Margins => new(Math.Max(0, (Left ?? 0) + Math.Min(0, FirstLine ?? 0)),
        Before ?? 0, Math.Max(0, Right ?? 0), After ?? 0);
    internal OfficeTextParagraphIndent Indent => new(Math.Max(0, FirstLine ?? 0), Math.Max(0, -(FirstLine ?? 0)));
    internal OfficeTextTabStops TabStops => new(Tabs ?? Array.Empty<OfficeTextTabStop>(), origin: -Margins.Left);

    internal PublisherParagraphStyle Inherit(PublisherParagraphStyle fallback) => new() {
        DefaultStyleIndex = DefaultStyleIndex ?? fallback.DefaultStyleIndex,
        Alignment = Alignment ?? fallback.Alignment,
        LineHeight = LineHeight ?? (LineHeightFactor.HasValue ? null : fallback.LineHeight),
        LineHeightFactor = LineHeightFactor ?? (LineHeight.HasValue ? null : fallback.LineHeightFactor),
        Before = Before ?? fallback.Before, After = After ?? fallback.After,
        Left = Left ?? fallback.Left, Right = Right ?? fallback.Right, FirstLine = FirstLine ?? fallback.FirstLine,
        Tabs = Tabs ?? fallback.Tabs, List = List ?? fallback.List, LabelSize = LabelSize ?? fallback.LabelSize,
        LabelFont = LabelFont ?? fallback.LabelFont, LabelTextPosition = LabelTextPosition ?? fallback.LabelTextPosition
    };
}

internal sealed record PublisherListStyle(uint NumberingType, uint Bullet);
