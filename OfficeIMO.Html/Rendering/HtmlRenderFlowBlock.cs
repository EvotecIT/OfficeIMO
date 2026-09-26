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
        double? unclampedHeight = null,
        IEnumerable<HtmlCssRunningStringAssignment>? runningStringAssignments = null,
        IEnumerable<HtmlInlineBreakProgress>? inlineBreakProgress = null,
        int inlineContinuationStart = 0,
        bool supportsInlineContinuationReflow = false,
        IEnumerable<HtmlRenderForcedBreak>? forcedBreaks = null,
        double layoutViewportWidth = double.NaN,
        double layoutViewportHeight = double.NaN,
        double leadingFlowAdjustment = 0D,
        HtmlCollapsedMargin? collapsibleMarginTopGroup = null,
        HtmlCollapsedMargin? collapsibleMarginBottomGroup = null,
        IEnumerable<HtmlRenderAvoidBreakRange>? avoidBreakRanges = null) {
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
        AvoidBreakRanges = new List<HtmlRenderAvoidBreakRange>(avoidBreakRanges ?? Array.Empty<HtmlRenderAvoidBreakRange>())
            .Where(range => range.Start >= -0.0001D && range.End <= height + 0.0001D && range.End > range.Start + 0.0001D)
            .ToList()
            .AsReadOnly();
        ForcedBreaks = new List<HtmlRenderForcedBreak>(forcedBreaks ?? Array.Empty<HtmlRenderForcedBreak>())
            .Where(item => item.Target != HtmlPageBreakTarget.None
                && !double.IsNaN(item.Offset)
                && !double.IsInfinity(item.Offset)
                && item.Offset >= -0.0001D
                && item.Offset <= height + 0.0001D)
            .OrderBy(item => item.Offset)
            .ToList()
            .AsReadOnly();
        var lineOffsets = new SortedSet<double>();
        bool hasImplicitFinalLine = false;
        if (lineBreakOffsets != null) {
            foreach (double offset in lineBreakOffsets) {
                if (offset <= 0D || double.IsNaN(offset) || double.IsInfinity(offset)) continue;
                if (offset < height - 0.0001D) lineOffsets.Add(offset);
                else if (offset <= height + 0.0001D) hasImplicitFinalLine = true;
            }
        }

        IReadOnlyList<double> resolvedLineOffsets = lineOffsets.ToList().AsReadOnly();
        int resolvedOrphans = Math.Max(1, orphans);
        int resolvedWidows = Math.Max(1, widows);
        var groups = new List<HtmlRenderLineBreakGroup>();
        if (lineBreakGroups != null) groups.AddRange(lineBreakGroups);
        if (groups.Count == 0 && resolvedLineOffsets.Count > 0) groups.Add(new HtmlRenderLineBreakGroup(resolvedLineOffsets, resolvedOrphans, resolvedWidows, hasImplicitFinalLine));
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
        CollapsibleMarginTopGroup = collapsibleMarginTopGroup ?? new HtmlCollapsedMargin(collapsibleMarginTop);
        CollapsibleMarginBottomGroup = collapsibleMarginBottomGroup ?? new HtmlCollapsedMargin(collapsibleMarginBottom);
        OwnerElement = ownerElement;
        CollapsesThrough = collapsesThrough;
        RunningStringAssignments = new List<HtmlCssRunningStringAssignment>(runningStringAssignments ?? Array.Empty<HtmlCssRunningStringAssignment>()).AsReadOnly();
        InlineBreakProgress = new List<HtmlInlineBreakProgress>(inlineBreakProgress ?? Array.Empty<HtmlInlineBreakProgress>()).AsReadOnly();
        InlineContinuationStart = Math.Max(0, inlineContinuationStart);
        SupportsInlineContinuationReflow = supportsInlineContinuationReflow;
        LayoutViewportWidth = layoutViewportWidth;
        LayoutViewportHeight = layoutViewportHeight;
        LeadingFlowAdjustment = leadingFlowAdjustment;
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
    internal IReadOnlyList<HtmlRenderAvoidBreakRange> AvoidBreakRanges { get; }
    internal IReadOnlyList<HtmlRenderForcedBreak> ForcedBreaks { get; }
    internal IReadOnlyList<HtmlRenderLineBreakGroup> LineBreakGroups { get; }
    internal IReadOnlyList<HtmlRenderContinuationGroup> ContinuationGroups { get; }
    internal IReadOnlyList<HtmlRenderTrailingGroup> TrailingGroups { get; }
    internal string? PageName { get; }
    internal int? StackingZIndex { get; }
    internal int StackingSourceOrder { get; }
    internal bool HasCollapsibleMargins { get; }
    internal double CollapsibleMarginTop { get; }
    internal double CollapsibleMarginBottom { get; }
    internal HtmlCollapsedMargin CollapsibleMarginTopGroup { get; }
    internal HtmlCollapsedMargin CollapsibleMarginBottomGroup { get; }
    internal IElement? OwnerElement { get; }
    internal bool CollapsesThrough { get; }
    internal IReadOnlyList<HtmlCssRunningStringAssignment> RunningStringAssignments { get; }
    internal IReadOnlyList<HtmlInlineBreakProgress> InlineBreakProgress { get; }
    internal int InlineContinuationStart { get; }
    internal bool SupportsInlineContinuationReflow { get; }
    internal double LayoutViewportWidth { get; }
    internal double LayoutViewportHeight { get; }
    internal double LeadingFlowAdjustment { get; }

    internal HtmlRenderFlowBlock WithLayoutViewport(double width, double height) =>
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
            hasCollapsibleMargins: HasCollapsibleMargins,
            collapsibleMarginTop: CollapsibleMarginTop,
            collapsibleMarginBottom: CollapsibleMarginBottom,
            ownerElement: OwnerElement,
            collapsesThrough: CollapsesThrough,
            unclampedHeight: UnclampedHeight,
            runningStringAssignments: RunningStringAssignments,
            inlineBreakProgress: InlineBreakProgress,
            inlineContinuationStart: InlineContinuationStart,
            supportsInlineContinuationReflow: SupportsInlineContinuationReflow,
            forcedBreaks: ForcedBreaks,
            layoutViewportWidth: width,
            layoutViewportHeight: height,
            leadingFlowAdjustment: LeadingFlowAdjustment,
            collapsibleMarginTopGroup: CollapsibleMarginTopGroup,
            collapsibleMarginBottomGroup: CollapsibleMarginBottomGroup,
            avoidBreakRanges: AvoidBreakRanges);

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
            unclampedHeight: UnclampedHeight,
            runningStringAssignments: RunningStringAssignments,
            inlineBreakProgress: InlineBreakProgress,
            inlineContinuationStart: InlineContinuationStart,
            supportsInlineContinuationReflow: SupportsInlineContinuationReflow,
            forcedBreaks: ForcedBreaks,
            layoutViewportWidth: LayoutViewportWidth,
            layoutViewportHeight: LayoutViewportHeight,
            leadingFlowAdjustment: LeadingFlowAdjustment,
            collapsibleMarginTopGroup: CollapsibleMarginTopGroup,
            collapsibleMarginBottomGroup: CollapsibleMarginBottomGroup,
            avoidBreakRanges: AvoidBreakRanges);

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
            unclampedHeight: UnclampedHeight,
            runningStringAssignments: RunningStringAssignments,
            inlineBreakProgress: InlineBreakProgress,
            inlineContinuationStart: InlineContinuationStart,
            supportsInlineContinuationReflow: SupportsInlineContinuationReflow,
            forcedBreaks: ForcedBreaks,
            layoutViewportWidth: LayoutViewportWidth,
            layoutViewportHeight: LayoutViewportHeight,
            leadingFlowAdjustment: LeadingFlowAdjustment,
            collapsibleMarginTopGroup: CollapsibleMarginTopGroup,
            collapsibleMarginBottomGroup: CollapsibleMarginBottomGroup,
            avoidBreakRanges: AvoidBreakRanges);

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
            unclampedHeight: UnclampedHeight,
            runningStringAssignments: RunningStringAssignments,
            inlineBreakProgress: InlineBreakProgress,
            inlineContinuationStart: InlineContinuationStart,
            supportsInlineContinuationReflow: SupportsInlineContinuationReflow,
            forcedBreaks: ForcedBreaks,
            layoutViewportWidth: LayoutViewportWidth,
            layoutViewportHeight: LayoutViewportHeight,
            leadingFlowAdjustment: LeadingFlowAdjustment,
            collapsibleMarginTopGroup: CollapsibleMarginTopGroup,
            collapsibleMarginBottomGroup: CollapsibleMarginBottomGroup,
            avoidBreakRanges: AvoidBreakRanges);

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
            unclampedHeight: adjustedUnclampedHeight,
            runningStringAssignments: RunningStringAssignments.Select(assignment => assignment.Translate(-adjustment)),
            inlineBreakProgress: InlineBreakProgress.Select(progress => new HtmlInlineBreakProgress(progress.Offset - adjustment, progress.LogicalCharacters, progress.OwnerElement, progress.IsBlockEntry, progress.PageStartDiscardableMargin)),
            inlineContinuationStart: InlineContinuationStart,
            supportsInlineContinuationReflow: SupportsInlineContinuationReflow,
            forcedBreaks: ForcedBreaks.Select(item => item.Translate(-adjustment)),
            layoutViewportWidth: LayoutViewportWidth,
            layoutViewportHeight: LayoutViewportHeight,
            leadingFlowAdjustment: LeadingFlowAdjustment + adjustment,
            collapsibleMarginTopGroup: CollapsibleMarginTopGroup,
            collapsibleMarginBottomGroup: CollapsibleMarginBottomGroup,
            avoidBreakRanges: AvoidBreakRanges.Select(range =>
                new HtmlRenderAvoidBreakRange(Math.Max(0D, range.Start - adjustment), range.End - adjustment)));
    }

    internal HtmlRenderFlowBlock WithCollapsibleMargins(double top, double bottom, IElement ownerElement,
        bool collapsesThrough = false, HtmlCollapsedMargin? topGroup = null, HtmlCollapsedMargin? bottomGroup = null) =>
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
            unclampedHeight: UnclampedHeight,
            runningStringAssignments: RunningStringAssignments,
            inlineBreakProgress: InlineBreakProgress,
            inlineContinuationStart: InlineContinuationStart,
            supportsInlineContinuationReflow: SupportsInlineContinuationReflow,
            forcedBreaks: ForcedBreaks,
            layoutViewportWidth: LayoutViewportWidth,
            layoutViewportHeight: LayoutViewportHeight,
            leadingFlowAdjustment: LeadingFlowAdjustment,
            collapsibleMarginTopGroup: topGroup,
            collapsibleMarginBottomGroup: bottomGroup,
            avoidBreakRanges: AvoidBreakRanges);

    internal HtmlRenderFlowBlock WithRunningStringAssignments(IEnumerable<HtmlCssRunningStringAssignment> assignments) =>
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
            hasCollapsibleMargins: HasCollapsibleMargins,
            collapsibleMarginTop: CollapsibleMarginTop,
            collapsibleMarginBottom: CollapsibleMarginBottom,
            ownerElement: OwnerElement,
            collapsesThrough: CollapsesThrough,
            unclampedHeight: UnclampedHeight,
            runningStringAssignments: assignments,
            inlineBreakProgress: InlineBreakProgress,
            inlineContinuationStart: InlineContinuationStart,
            supportsInlineContinuationReflow: SupportsInlineContinuationReflow,
            forcedBreaks: ForcedBreaks,
            layoutViewportWidth: LayoutViewportWidth,
            layoutViewportHeight: LayoutViewportHeight,
            leadingFlowAdjustment: LeadingFlowAdjustment,
            collapsibleMarginTopGroup: CollapsibleMarginTopGroup,
            collapsibleMarginBottomGroup: CollapsibleMarginBottomGroup,
            avoidBreakRanges: AvoidBreakRanges);

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
            unclampedHeight: adjustedUnclampedHeight,
            runningStringAssignments: RunningStringAssignments.Where(assignment => assignment.Offset <= adjustedHeight + 0.0001D),
            inlineBreakProgress: InlineBreakProgress.Where(progress => progress.Offset <= adjustedHeight + 0.0001D),
            inlineContinuationStart: InlineContinuationStart,
            supportsInlineContinuationReflow: SupportsInlineContinuationReflow,
            forcedBreaks: ForcedBreaks.Where(item => item.Offset <= adjustedHeight + 0.0001D),
            layoutViewportWidth: LayoutViewportWidth,
            layoutViewportHeight: LayoutViewportHeight,
            leadingFlowAdjustment: LeadingFlowAdjustment,
            collapsibleMarginTopGroup: CollapsibleMarginTopGroup,
            collapsibleMarginBottomGroup: CollapsibleMarginBottomGroup,
            avoidBreakRanges: AvoidBreakRanges.Select(range => range.WithEnd(Math.Min(range.End, adjustedHeight))));
    }
}

internal readonly record struct HtmlRenderAvoidBreakRange(double Start, double End) {
    internal HtmlRenderAvoidBreakRange Translate(double offset) => new HtmlRenderAvoidBreakRange(Start + offset, End + offset);
    internal HtmlRenderAvoidBreakRange WithEnd(double end) => new HtmlRenderAvoidBreakRange(Start, end);
}

internal sealed class HtmlRenderForcedBreak {
    internal HtmlRenderForcedBreak(double offset, HtmlPageBreakTarget target, string? pageName = null, bool changesPageName = false) {
        Offset = offset;
        Target = target;
        PageName = pageName;
        ChangesPageName = changesPageName;
    }

    internal double Offset { get; }
    internal HtmlPageBreakTarget Target { get; }
    internal string? PageName { get; }
    internal bool ChangesPageName { get; }

    internal HtmlRenderForcedBreak Translate(double offset) => new HtmlRenderForcedBreak(Offset + offset, Target, PageName, ChangesPageName);
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
    internal HtmlRenderLineBreakGroup(IEnumerable<double> offsets, int orphans, int widows, bool hasImplicitFinalLine = false) {
        Offsets = new SortedSet<double>(offsets).ToList().AsReadOnly();
        Orphans = Math.Max(1, orphans);
        Widows = Math.Max(1, widows);
        HasImplicitFinalLine = hasImplicitFinalLine;
    }

    internal IReadOnlyList<double> Offsets { get; }
    internal int Orphans { get; }
    internal int Widows { get; }
    internal bool HasImplicitFinalLine { get; }

    internal HtmlRenderLineBreakGroup Translate(double offset) =>
        new HtmlRenderLineBreakGroup(Offsets.Select(value => value + offset), Orphans, Widows, HasImplicitFinalLine);
}

internal sealed class HtmlInlineRun {
    internal HtmlInlineRun(
        IEnumerable<HtmlCssRunningStringAssignment> runningElementAssignments,
        HtmlRenderBoxStyle style,
        string source) {
        RunningElementAssignments = new List<HtmlCssRunningStringAssignment>(runningElementAssignments).AsReadOnly();
        RunningElementAssignment = RunningElementAssignments.FirstOrDefault();
        Text = string.Empty;
        LogicalText = string.Empty;
        Style = style;
        Source = source;
        SemanticRole = style.SemanticRole;
    }

    internal HtmlInlineRun(
        IElement runningStringElement,
        HtmlRenderBoxStyle style,
        string source) {
        RunningStringElement = runningStringElement;
        Text = string.Empty;
        LogicalText = string.Empty;
        Style = style;
        Source = source;
        SemanticRole = style.SemanticRole;
    }

    internal HtmlInlineRun(
        string text,
        HtmlRenderBoxStyle style,
        string? linkUri,
        string source,
        double paintOffsetX = 0D,
        double paintOffsetY = 0D,
        IElement? ownerElement = null,
        IElement? positionedMarkerElement = null,
        string? logicalText = null,
        bool textTransformPending = false,
        string? leaderPattern = null) {
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
        TextTransformPending = textTransformPending;
        LeaderPattern = leaderPattern;
    }

    internal HtmlInlineRun(
        HtmlRenderFlowBlock atomicBlock,
        HtmlRenderBoxStyle style,
        string? linkUri,
        string source,
        double paintOffsetX = 0D,
        double paintOffsetY = 0D,
        IElement? ownerElement = null,
        bool isReplacedImage = false,
        double? atomicBaseline = null,
        bool isBookmarkMarker = false) {
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
        AtomicBaseline = atomicBaseline;
        IsBookmarkMarker = isBookmarkMarker;
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

    internal string Text { get; private set; }
    internal string LogicalText { get; private set; }
    internal HtmlRenderFlowBlock? AtomicBlock { get; }
    internal HtmlRenderFlowBlock? FloatingBlock { get; }
    internal HtmlRenderBoxStyle Style { get; }
    internal string? LinkUri { get; }
    internal string Source { get; }
    internal double PaintOffsetX { get; }
    internal double PaintOffsetY { get; }
    internal IElement? OwnerElement { get; }
    internal IElement? PositionedMarkerElement { get; }
    internal IElement? RunningStringElement { get; }
    internal HtmlCssRunningStringAssignment? RunningElementAssignment { get; }
    internal IReadOnlyList<HtmlCssRunningStringAssignment> RunningElementAssignments { get; } = Array.Empty<HtmlCssRunningStringAssignment>();
    internal bool IsReplacedImage { get; }
    internal double? AtomicBaseline { get; }
    internal bool IsBookmarkMarker { get; }
    internal string SemanticRole { get; private set; }
    internal int? SemanticNodeId { get; private set; }
    internal int? SemanticFragmentOrder { get; private set; }
    internal int? LogicalTextOrder { get; private set; }
    internal HtmlRenderSemanticGroupRole? InlineSemanticGroupRole { get; private set; }
    internal string? InlineSemanticGroupKey { get; private set; }
    internal string? BookmarkAnchorText { get; private set; }
    internal string FloatSide { get; } = "none";
    internal string ClearSide { get; } = "none";
    internal bool TextTransformPending { get; private set; }
    internal string? LeaderPattern { get; }
    internal bool IsFirstLetter { get; private set; }

    internal void CompleteTextTransform(string text) {
        Text = text;
        LogicalText = text;
        TextTransformPending = false;
    }

    internal void AssignSemanticNode(string role, int nodeId, string? bookmarkAnchorText = null, int? semanticFragmentOrder = null) {
        SemanticNodeId = nodeId;
        SemanticFragmentOrder = semanticFragmentOrder;
        BookmarkAnchorText = bookmarkAnchorText;
        if (!SemanticRole.StartsWith("generated-", StringComparison.Ordinal)) {
            SemanticRole = role;
        }
    }

    internal void AssignLogicalTextOrder(int order) => LogicalTextOrder = order;

    internal void AssignInlineSemanticGroup(HtmlRenderSemanticGroupRole role, string structureElementKey) {
        InlineSemanticGroupRole = role;
        InlineSemanticGroupKey = structureElementKey;
    }

    internal HtmlInlineRun CloneText(string text, string logicalText, HtmlRenderBoxStyle style, bool isFirstLetter = false) {
        var clone = new HtmlInlineRun(
            text,
            style,
            LinkUri,
            Source,
            PaintOffsetX,
            PaintOffsetY,
            OwnerElement,
            PositionedMarkerElement,
            logicalText,
            TextTransformPending,
            LeaderPattern) {
            SemanticRole = SemanticRole,
            SemanticNodeId = SemanticNodeId,
            SemanticFragmentOrder = SemanticFragmentOrder,
            LogicalTextOrder = LogicalTextOrder,
            InlineSemanticGroupRole = InlineSemanticGroupRole,
            InlineSemanticGroupKey = InlineSemanticGroupKey,
            BookmarkAnchorText = BookmarkAnchorText,
            IsFirstLetter = isFirstLetter
        };
        return clone;
    }
}

internal sealed class HtmlInlineLayout {
    internal HtmlInlineLayout(
        IEnumerable<HtmlRenderVisual> visuals,
        double height,
        IEnumerable<double>? breakOffsets = null,
        IEnumerable<HtmlCssRunningStringAssignment>? runningStringAssignments = null,
        IEnumerable<HtmlInlineBreakProgress>? breakProgress = null,
        bool supportsContinuationReflow = false,
        double? normalFlowHeight = null,
        IEnumerable<double>? lineBreakOffsets = null) {
        Visuals = new List<HtmlRenderVisual>(visuals);
        Height = height;
        NormalFlowHeight = normalFlowHeight ?? height;
        BreakOffsets = new List<double>(breakOffsets ?? Array.Empty<double>()).AsReadOnly();
        LineBreakOffsets = new List<double>(lineBreakOffsets ?? BreakOffsets).AsReadOnly();
        RunningStringAssignments = new List<HtmlCssRunningStringAssignment>(
            runningStringAssignments ?? Array.Empty<HtmlCssRunningStringAssignment>()).AsReadOnly();
        BreakProgress = new List<HtmlInlineBreakProgress>(breakProgress ?? Array.Empty<HtmlInlineBreakProgress>()).AsReadOnly();
        SupportsContinuationReflow = supportsContinuationReflow;
    }

    internal IReadOnlyList<HtmlRenderVisual> Visuals { get; }
    internal double Height { get; }
    internal double NormalFlowHeight { get; }
    /// <summary>Page-break candidates; may exclude line ends inside a floated box.</summary>
    internal IReadOnlyList<double> BreakOffsets { get; }
    /// <summary>All line ends used to count widows and orphans, including lines beside floats.</summary>
    internal IReadOnlyList<double> LineBreakOffsets { get; }
    internal IReadOnlyList<HtmlCssRunningStringAssignment> RunningStringAssignments { get; }
    internal IReadOnlyList<HtmlInlineBreakProgress> BreakProgress { get; }
    internal bool SupportsContinuationReflow { get; }
}

internal readonly struct HtmlInlineBreakProgress {
    internal HtmlInlineBreakProgress(double offset, int logicalCharacters, IElement? ownerElement = null, bool isBlockEntry = false, double pageStartDiscardableMargin = 0D) {
        Offset = offset;
        LogicalCharacters = logicalCharacters;
        OwnerElement = ownerElement;
        IsBlockEntry = isBlockEntry;
        PageStartDiscardableMargin = pageStartDiscardableMargin;
    }

    internal double Offset { get; }
    internal int LogicalCharacters { get; }
    internal IElement? OwnerElement { get; }
    internal bool IsBlockEntry { get; }
    internal double PageStartDiscardableMargin { get; }
}
