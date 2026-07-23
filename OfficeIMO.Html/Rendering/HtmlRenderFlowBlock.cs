using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed class HtmlRenderFlowBlock {
    internal HtmlRenderFlowBlock(
        double width,
        double height,
        IEnumerable<HtmlRenderVisual> visuals,
        HtmlPageBreakTarget breakBefore,
        HtmlPageBreakTarget breakAfter,
        bool avoidBreakInside,
        string source,
        IEnumerable<double>? breakOffsets = null,
        IEnumerable<double>? lineBreakOffsets = null,
        int orphans = 2,
        int widows = 2,
        IEnumerable<HtmlRenderLineBreakGroup>? lineBreakGroups = null,
        IEnumerable<HtmlRenderContinuationGroup>? continuationGroups = null,
        IEnumerable<HtmlRenderTrailingGroup>? trailingGroups = null,
        IEnumerable<HtmlRenderVisual>? continuationVisuals = null,
        double continuationHeight = 0D,
        double continuationStartsAfter = 0D,
        string? pageName = null,
        int? stackingZIndex = null,
        int stackingSourceOrder = 0,
        bool hasCollapsibleMargins = false,
        double collapsibleMarginTop = 0D,
        double collapsibleMarginBottom = 0D,
        IElement? ownerElement = null,
        bool collapsesThrough = false,
        double? unclampedHeight = null) {
        Width = width;
        Height = height;
        UnclampedHeight = unclampedHeight.HasValue && !double.IsNaN(unclampedHeight.Value) && !double.IsInfinity(unclampedHeight.Value)
            ? unclampedHeight.Value
            : height;
        Visuals = new List<HtmlRenderVisual>(visuals);
        BreakBefore = breakBefore;
        BreakAfter = breakAfter;
        AvoidBreakInside = avoidBreakInside;
        Source = source;
        var offsets = new SortedSet<double> { 0D, height };
        if (breakOffsets != null) {
            foreach (double offset in breakOffsets) {
                if (offset > 0D && offset < height && !double.IsNaN(offset) && !double.IsInfinity(offset)) offsets.Add(offset);
            }
        }

        BreakOffsets = offsets.ToList().AsReadOnly();
        var lineOffsets = new SortedSet<double>();
        if (lineBreakOffsets != null) {
            foreach (double offset in lineBreakOffsets) {
                if (offset > 0D && offset < height && !double.IsNaN(offset) && !double.IsInfinity(offset)) lineOffsets.Add(offset);
            }
        }

        IReadOnlyList<double> resolvedLineOffsets = lineOffsets.ToList().AsReadOnly();
        int resolvedOrphans = Math.Max(1, orphans);
        int resolvedWidows = Math.Max(1, widows);
        var groups = new List<HtmlRenderLineBreakGroup>();
        if (lineBreakGroups != null) groups.AddRange(lineBreakGroups);
        if (groups.Count == 0 && resolvedLineOffsets.Count > 0) groups.Add(new HtmlRenderLineBreakGroup(resolvedLineOffsets, resolvedOrphans, resolvedWidows));
        LineBreakGroups = groups.AsReadOnly();
        IReadOnlyList<HtmlRenderVisual> repeatedVisuals = new List<HtmlRenderVisual>(continuationVisuals ?? Array.Empty<HtmlRenderVisual>()).AsReadOnly();
        double repeatedHeight = Math.Max(0D, continuationHeight);
        double repeatedStartsAfter = Math.Max(0D, continuationStartsAfter);
        var repeatedGroups = new List<HtmlRenderContinuationGroup>(continuationGroups ?? Array.Empty<HtmlRenderContinuationGroup>());
        if (repeatedGroups.Count == 0 && repeatedVisuals.Count > 0 && repeatedHeight > 0D) {
            repeatedGroups.Add(new HtmlRenderContinuationGroup(repeatedStartsAfter, height, repeatedHeight, repeatedVisuals));
        }

        ContinuationGroups = repeatedGroups.AsReadOnly();
        TrailingGroups = new List<HtmlRenderTrailingGroup>(trailingGroups ?? Array.Empty<HtmlRenderTrailingGroup>()).AsReadOnly();
        PageName = pageName == null || string.IsNullOrWhiteSpace(pageName) ? null : pageName.Trim();
        StackingZIndex = stackingZIndex;
        StackingSourceOrder = stackingSourceOrder;
        HasCollapsibleMargins = hasCollapsibleMargins;
        CollapsibleMarginTop = collapsibleMarginTop;
        CollapsibleMarginBottom = collapsibleMarginBottom;
        OwnerElement = ownerElement;
        CollapsesThrough = collapsesThrough;
    }

    internal double Width { get; }
    internal double Height { get; }
    internal double UnclampedHeight { get; }
    internal IReadOnlyList<HtmlRenderVisual> Visuals { get; }
    internal HtmlPageBreakTarget BreakBefore { get; }
    internal HtmlPageBreakTarget BreakAfter { get; }
    internal bool AvoidBreakInside { get; }
    internal string Source { get; }
    internal IReadOnlyList<double> BreakOffsets { get; }
    internal IReadOnlyList<HtmlRenderLineBreakGroup> LineBreakGroups { get; }
    internal IReadOnlyList<HtmlRenderContinuationGroup> ContinuationGroups { get; }
    internal IReadOnlyList<HtmlRenderTrailingGroup> TrailingGroups { get; }
    internal string? PageName { get; }
    internal int? StackingZIndex { get; }
    internal int StackingSourceOrder { get; }
    internal bool HasCollapsibleMargins { get; }
    internal double CollapsibleMarginTop { get; }
    internal double CollapsibleMarginBottom { get; }
    internal IElement? OwnerElement { get; }
    internal bool CollapsesThrough { get; }

    internal HtmlRenderFlowBlock TranslatePaint(double offsetX, double offsetY) =>
        new HtmlRenderFlowBlock(
            Width,
            Height,
            Visuals.Select(visual => visual.TranslatePaint(offsetX, offsetY, visual.PaintOrder)),
            BreakBefore,
            BreakAfter,
            AvoidBreakInside,
            Source,
            BreakOffsets,
            lineBreakGroups: LineBreakGroups,
            continuationGroups: ContinuationGroups.Select(group => group.TranslatePaint(offsetX, offsetY)),
            trailingGroups: TrailingGroups.Select(group => group.TranslatePaint(offsetX, offsetY)),
            pageName: PageName,
            stackingZIndex: StackingZIndex,
            stackingSourceOrder: StackingSourceOrder,
            hasCollapsibleMargins: HasCollapsibleMargins,
            collapsibleMarginTop: CollapsibleMarginTop,
            collapsibleMarginBottom: CollapsibleMarginBottom,
            ownerElement: OwnerElement,
            collapsesThrough: CollapsesThrough,
            unclampedHeight: UnclampedHeight);

    internal HtmlRenderFlowBlock WithStacking(int zIndex, int sourceOrder) =>
        new HtmlRenderFlowBlock(
            Width,
            Height,
            Visuals,
            BreakBefore,
            BreakAfter,
            AvoidBreakInside,
            Source,
            BreakOffsets,
            lineBreakGroups: LineBreakGroups,
            continuationGroups: ContinuationGroups,
            trailingGroups: TrailingGroups,
            pageName: PageName,
            stackingZIndex: zIndex,
            stackingSourceOrder: sourceOrder,
            hasCollapsibleMargins: HasCollapsibleMargins,
            collapsibleMarginTop: CollapsibleMarginTop,
            collapsibleMarginBottom: CollapsibleMarginBottom,
            ownerElement: OwnerElement,
            collapsesThrough: CollapsesThrough,
            unclampedHeight: UnclampedHeight);

    internal HtmlRenderFlowBlock WithVisuals(IEnumerable<HtmlRenderVisual> visuals) =>
        new HtmlRenderFlowBlock(
            Width,
            Height,
            visuals,
            BreakBefore,
            BreakAfter,
            AvoidBreakInside,
            Source,
            BreakOffsets,
            lineBreakGroups: LineBreakGroups,
            continuationGroups: ContinuationGroups,
            trailingGroups: TrailingGroups,
            pageName: PageName,
            stackingZIndex: StackingZIndex,
            stackingSourceOrder: StackingSourceOrder,
            hasCollapsibleMargins: HasCollapsibleMargins,
            collapsibleMarginTop: CollapsibleMarginTop,
            collapsibleMarginBottom: CollapsibleMarginBottom,
            ownerElement: OwnerElement,
            collapsesThrough: CollapsesThrough,
            unclampedHeight: UnclampedHeight);

    internal HtmlRenderFlowBlock AdjustLeadingFlowSpace(double adjustment) {
        if (Math.Abs(adjustment) <= 0.0001D) return this;
        double adjustedUnclampedHeight = UnclampedHeight - adjustment;
        double adjustedHeight = Math.Max(0.01D, adjustedUnclampedHeight);
        return new HtmlRenderFlowBlock(
            Width,
            adjustedHeight,
            Visuals.Select((visual, index) => visual.Translate(0D, -adjustment, index)),
            BreakBefore,
            BreakAfter,
            AvoidBreakInside,
            Source,
            BreakOffsets.Select(offset => offset - adjustment),
            lineBreakGroups: LineBreakGroups.Select(group => group.Translate(-adjustment)),
            continuationGroups: ContinuationGroups.Select(group => group.Translate(0D, -adjustment)),
            trailingGroups: TrailingGroups.Select(group => group.Translate(0D, -adjustment)),
            pageName: PageName,
            stackingZIndex: StackingZIndex,
            stackingSourceOrder: StackingSourceOrder,
            hasCollapsibleMargins: HasCollapsibleMargins,
            collapsibleMarginTop: CollapsibleMarginTop,
            collapsibleMarginBottom: CollapsibleMarginBottom,
            ownerElement: OwnerElement,
            collapsesThrough: CollapsesThrough,
            unclampedHeight: adjustedUnclampedHeight);
    }

    internal HtmlRenderFlowBlock WithCollapsibleMargins(double top, double bottom, IElement ownerElement, bool collapsesThrough = false) =>
        new HtmlRenderFlowBlock(
            Width,
            Height,
            Visuals,
            BreakBefore,
            BreakAfter,
            AvoidBreakInside,
            Source,
            BreakOffsets,
            lineBreakGroups: LineBreakGroups,
            continuationGroups: ContinuationGroups,
            trailingGroups: TrailingGroups,
            pageName: PageName,
            stackingZIndex: StackingZIndex,
            stackingSourceOrder: StackingSourceOrder,
            hasCollapsibleMargins: true,
            collapsibleMarginTop: top,
            collapsibleMarginBottom: bottom,
            ownerElement: ownerElement,
            collapsesThrough: collapsesThrough,
            unclampedHeight: UnclampedHeight);

    internal HtmlRenderFlowBlock AdjustTrailingFlowSpace(double adjustment) {
        if (Math.Abs(adjustment) <= 0.0001D) return this;
        double adjustedUnclampedHeight = UnclampedHeight - adjustment;
        double adjustedHeight = Math.Max(0.01D, adjustedUnclampedHeight);
        return new HtmlRenderFlowBlock(
            Width,
            adjustedHeight,
            Visuals,
            BreakBefore,
            BreakAfter,
            AvoidBreakInside,
            Source,
            BreakOffsets.Where(offset => offset <= adjustedHeight + 0.0001D),
            lineBreakGroups: LineBreakGroups,
            continuationGroups: ContinuationGroups,
            trailingGroups: TrailingGroups,
            pageName: PageName,
            stackingZIndex: StackingZIndex,
            stackingSourceOrder: StackingSourceOrder,
            hasCollapsibleMargins: HasCollapsibleMargins,
            collapsibleMarginTop: CollapsibleMarginTop,
            collapsibleMarginBottom: CollapsibleMarginBottom,
            ownerElement: OwnerElement,
            collapsesThrough: CollapsesThrough,
            unclampedHeight: adjustedUnclampedHeight);
    }
}

internal sealed class HtmlRenderContinuationGroup {
    internal HtmlRenderContinuationGroup(double startsAfter, double endsAt, double height, IEnumerable<HtmlRenderVisual> visuals) {
        StartsAfter = startsAfter;
        EndsAt = endsAt;
        Height = height;
        Visuals = new List<HtmlRenderVisual>(visuals).AsReadOnly();
    }

    internal double StartsAfter { get; }
    internal double EndsAt { get; }
    internal double Height { get; }
    internal IReadOnlyList<HtmlRenderVisual> Visuals { get; }

    internal bool AppliesAt(double offset) => offset >= StartsAfter - 0.0001D && offset < EndsAt - 0.0001D;

    internal HtmlRenderContinuationGroup Translate(double offsetX, double offsetY) =>
        new HtmlRenderContinuationGroup(
            StartsAfter + offsetY,
            EndsAt + offsetY,
            Height,
            Visuals.Select((visual, index) => visual.Translate(offsetX, 0D, index)));

    internal HtmlRenderContinuationGroup TranslatePaint(double offsetX, double offsetY) =>
        new HtmlRenderContinuationGroup(
            StartsAfter,
            EndsAt,
            Height,
            Visuals.Select(visual => visual.TranslatePaint(offsetX, offsetY, visual.PaintOrder)));
}

internal sealed class HtmlRenderTrailingGroup {
    internal HtmlRenderTrailingGroup(double startsAt, double contentEndsAt, double sourceEndsAt, double height, IEnumerable<HtmlRenderVisual> visuals) {
        StartsAt = startsAt;
        ContentEndsAt = contentEndsAt;
        SourceEndsAt = sourceEndsAt;
        Height = height;
        Visuals = new List<HtmlRenderVisual>(visuals).AsReadOnly();
    }

    internal double StartsAt { get; }
    internal double ContentEndsAt { get; }
    internal double SourceEndsAt { get; }
    internal double Height { get; }
    internal IReadOnlyList<HtmlRenderVisual> Visuals { get; }

    internal bool AppliesAt(double offset) => offset >= StartsAt - 0.0001D && offset < ContentEndsAt - 0.0001D;

    internal HtmlRenderTrailingGroup Translate(double offsetX, double offsetY, double? sourceEndsAt = null) {
        double translatedSourceEnd = SourceEndsAt + offsetY;
        double resolvedSourceEnd = sourceEndsAt ?? translatedSourceEnd;
        return new HtmlRenderTrailingGroup(
            StartsAt + offsetY,
            ContentEndsAt + offsetY,
            resolvedSourceEnd,
            Height + Math.Max(0D, resolvedSourceEnd - translatedSourceEnd),
            Visuals.Select((visual, index) => visual.Translate(offsetX, 0D, index)));
    }

    internal HtmlRenderTrailingGroup TranslatePaint(double offsetX, double offsetY) =>
        new HtmlRenderTrailingGroup(
            StartsAt,
            ContentEndsAt,
            SourceEndsAt,
            Height,
            Visuals.Select(visual => visual.TranslatePaint(offsetX, offsetY, visual.PaintOrder)));
}

internal sealed class HtmlRenderLineBreakGroup {
    internal HtmlRenderLineBreakGroup(IEnumerable<double> offsets, int orphans, int widows) {
        Offsets = new SortedSet<double>(offsets).ToList().AsReadOnly();
        Orphans = Math.Max(1, orphans);
        Widows = Math.Max(1, widows);
    }

    internal IReadOnlyList<double> Offsets { get; }
    internal int Orphans { get; }
    internal int Widows { get; }

    internal HtmlRenderLineBreakGroup Translate(double offset) =>
        new HtmlRenderLineBreakGroup(Offsets.Select(value => value + offset), Orphans, Widows);
}

internal sealed class HtmlInlineRun {
    internal HtmlInlineRun(
        string text,
        HtmlRenderBoxStyle style,
        string? linkUri,
        string source,
        double paintOffsetX = 0D,
        double paintOffsetY = 0D,
        IElement? ownerElement = null,
        IElement? positionedMarkerElement = null,
        string? logicalText = null) {
        Text = text;
        LogicalText = logicalText ?? text;
        Style = style;
        LinkUri = linkUri;
        Source = source;
        PaintOffsetX = paintOffsetX;
        PaintOffsetY = paintOffsetY;
        OwnerElement = ownerElement;
        PositionedMarkerElement = positionedMarkerElement;
        SemanticRole = style.SemanticRole;
    }

    internal HtmlInlineRun(
        HtmlRenderFlowBlock atomicBlock,
        HtmlRenderBoxStyle style,
        string? linkUri,
        string source,
        double paintOffsetX = 0D,
        double paintOffsetY = 0D,
        IElement? ownerElement = null,
        bool isReplacedImage = false) {
        AtomicBlock = atomicBlock;
        Text = string.Empty;
        LogicalText = string.Empty;
        Style = style;
        LinkUri = linkUri;
        Source = source;
        PaintOffsetX = paintOffsetX;
        PaintOffsetY = paintOffsetY;
        OwnerElement = ownerElement;
        IsReplacedImage = isReplacedImage;
        SemanticRole = style.SemanticRole;
    }

    internal HtmlInlineRun(
        HtmlRenderFlowBlock floatingBlock,
        HtmlRenderBoxStyle style,
        string? linkUri,
        string source,
        string floatSide,
        string clearSide,
        IElement ownerElement) {
        FloatingBlock = floatingBlock;
        Text = string.Empty;
        LogicalText = string.Empty;
        Style = style;
        LinkUri = linkUri;
        Source = source;
        FloatSide = floatSide;
        ClearSide = clearSide;
        OwnerElement = ownerElement;
        SemanticRole = style.SemanticRole;
    }

    internal string Text { get; }
    internal string LogicalText { get; }
    internal HtmlRenderFlowBlock? AtomicBlock { get; }
    internal HtmlRenderFlowBlock? FloatingBlock { get; }
    internal HtmlRenderBoxStyle Style { get; }
    internal string? LinkUri { get; }
    internal string Source { get; }
    internal double PaintOffsetX { get; }
    internal double PaintOffsetY { get; }
    internal IElement? OwnerElement { get; }
    internal IElement? PositionedMarkerElement { get; }
    internal bool IsReplacedImage { get; }
    internal string SemanticRole { get; private set; }
    internal int? SemanticNodeId { get; private set; }
    internal string FloatSide { get; } = "none";
    internal string ClearSide { get; } = "none";

    internal void AssignSemanticNode(string role, int nodeId) {
        SemanticNodeId = nodeId;
        if (!SemanticRole.StartsWith("generated-", StringComparison.Ordinal)) {
            SemanticRole = role;
        }
    }
}

internal sealed class HtmlInlineLayout {
    internal HtmlInlineLayout(IEnumerable<HtmlRenderVisual> visuals, double height, IEnumerable<double>? breakOffsets = null) {
        Visuals = new List<HtmlRenderVisual>(visuals);
        Height = height;
        BreakOffsets = new List<double>(breakOffsets ?? Array.Empty<double>()).AsReadOnly();
    }

    internal IReadOnlyList<HtmlRenderVisual> Visuals { get; }
    internal double Height { get; }
    internal IReadOnlyList<double> BreakOffsets { get; }
}
