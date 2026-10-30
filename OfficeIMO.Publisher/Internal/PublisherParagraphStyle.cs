using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

// Nullable native values distinguish an explicit override from an inherited value.
internal sealed class PublisherParagraphStyle {
    public uint? DefaultStyleIndex { get; internal set; }
    public OfficeTextAlignment? Alignment { get; internal set; }
    public double? LineHeight { get; internal set; }
    public double? LineHeightFactor { get; internal set; }
    public double? Before { get; internal set; }
    public double? After { get; internal set; }
    public double? Left { get; internal set; }
    public double? Right { get; internal set; }
    public double? FirstLine { get; internal set; }
    public IReadOnlyList<OfficeTextTabStop>? Tabs { get; internal set; }
    public PublisherListStyle? List { get; internal set; }
    public double? LabelSize { get; internal set; }
    public uint? LabelFont { get; internal set; }
    public double? LabelTextPosition { get; internal set; }

    internal OfficeTextPadding Margins => new(Math.Max(0, (Left ?? 0) + Math.Min(0, FirstLine ?? 0)),
        Before ?? 0, Math.Max(0, Right ?? 0), After ?? 0);
    internal OfficeTextParagraphIndent Indent => new(Math.Max(0, FirstLine ?? 0), Math.Max(0, -(FirstLine ?? 0)));
    internal OfficeTextTabStops TabStops {
        get {
            var tabs = new OfficeTextTabStops(Tabs ?? Array.Empty<OfficeTextTabStop>(), origin: -Margins.Left);
            return Alignment is OfficeTextAlignment.Center or OfficeTextAlignment.Right ? tabs.WithParagraphAlignment() : tabs;
        }
    }

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
