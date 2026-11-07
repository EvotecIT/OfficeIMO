namespace OfficeIMO.Html;

internal sealed class HtmlInlineLayout {
    internal HtmlInlineLayout(
        IEnumerable<HtmlRenderVisual> visuals,
        double height,
        IEnumerable<double>? breakOffsets = null,
        IEnumerable<HtmlCssRunningStringAssignment>? runningStringAssignments = null,
        IEnumerable<HtmlInlineBreakProgress>? breakProgress = null,
        bool supportsContinuationReflow = false,
        double? normalFlowHeight = null,
        IEnumerable<double>? lineBreakOffsets = null,
        IEnumerable<HtmlFloatExclusion>? floatExclusions = null,
        HtmlRenderFlowBlock? interruptedFlow = null,
        IEnumerable<HtmlRenderForcedBreak>? forcedBreaks = null,
        IEnumerable<HtmlRenderLineBreakGroup>? lineBreakGroups = null,
        IEnumerable<HtmlRenderContinuationGroup>? continuationGroups = null,
        IEnumerable<HtmlRenderTrailingGroup>? trailingGroups = null,
        double? pagedPaintExtent = null,
        bool? hasInFlowLineBoxes = null,
        IEnumerable<HtmlRenderAvoidBreakRange>? avoidBreakRanges = null,
        IEnumerable<HtmlRenderBreakRange>? safeBreakRanges = null) {
        Visuals = new List<HtmlRenderVisual>(visuals);
        Height = height;
        NormalFlowHeight = normalFlowHeight ?? height;
        HasInFlowLineBoxes = hasInFlowLineBoxes;
        PagedPaintExtent = Math.Max(height, pagedPaintExtent ?? height);
        BreakOffsets = new List<double>(breakOffsets ?? Array.Empty<double>()).AsReadOnly();
        SafeBreakRanges = HtmlRenderBreakRange.Normalize(safeBreakRanges ?? Array.Empty<HtmlRenderBreakRange>(), PagedPaintExtent);
        LineBreakOffsets = new List<double>(lineBreakOffsets ?? BreakOffsets).AsReadOnly();
        RunningStringAssignments = SnapshotOptionalMetadata(runningStringAssignments);
        BreakProgress = new List<HtmlInlineBreakProgress>(breakProgress ?? Array.Empty<HtmlInlineBreakProgress>()).AsReadOnly();
        ForcedBreaks = forcedBreaks == null ? Array.Empty<HtmlRenderForcedBreak>() : new List<HtmlRenderForcedBreak>(forcedBreaks).AsReadOnly();
        LineBreakGroups = lineBreakGroups == null ? Array.Empty<HtmlRenderLineBreakGroup>() : new List<HtmlRenderLineBreakGroup>(lineBreakGroups).AsReadOnly();
        AvoidBreakRanges = avoidBreakRanges == null ? Array.Empty<HtmlRenderAvoidBreakRange>() : new List<HtmlRenderAvoidBreakRange>(avoidBreakRanges).AsReadOnly();
        ContinuationGroups = continuationGroups == null ? Array.Empty<HtmlRenderContinuationGroup>() : new List<HtmlRenderContinuationGroup>(continuationGroups).AsReadOnly();
        TrailingGroups = trailingGroups == null ? Array.Empty<HtmlRenderTrailingGroup>() : new List<HtmlRenderTrailingGroup>(trailingGroups).AsReadOnly();
        SupportsContinuationReflow = supportsContinuationReflow;
        InterruptedFlow = interruptedFlow;
        FloatExclusions = SnapshotOptionalMetadata(floatExclusions);
    }

    // Keep nonempty metadata detached; empty deferred sequences need no list or wrapper.
    private static IReadOnlyList<T> SnapshotOptionalMetadata<T>(IEnumerable<T>? source) {
        if (source == null || source is ICollection<T> { Count: 0 }) return Array.Empty<T>();
        if (source is ICollection<T>) return new List<T>(source).AsReadOnly();
        List<T>? snapshot = null;
        foreach (T item in source) (snapshot ??= new List<T>()).Add(item);
        return snapshot == null ? Array.Empty<T>() : snapshot.AsReadOnly();
    }

    internal IReadOnlyList<HtmlRenderVisual> Visuals { get; }
    internal double Height { get; }
    internal double NormalFlowHeight { get; }
    // Unknown on specialized layouts; ordinary line construction supplies the fact
    // before fragmentation filters or paint visibility can discard its evidence.
    internal bool? HasInFlowLineBoxes { get; }
    // Shared float contexts can paint beyond this chunk's normal-flow height.
    internal double PagedPaintExtent { get; }
    /// <summary>Page-break candidates; may exclude line ends inside a floated box.</summary>
    internal IReadOnlyList<double> BreakOffsets { get; }
    /// <summary>Closed source-coordinate intervals that can supply page-break candidates.</summary>
    internal IReadOnlyList<HtmlRenderBreakRange> SafeBreakRanges { get; }
    /// <summary>All line ends used to count widows and orphans, including lines beside floats.</summary>
    internal IReadOnlyList<double> LineBreakOffsets { get; }
    internal IReadOnlyList<HtmlCssRunningStringAssignment> RunningStringAssignments { get; }
    internal IReadOnlyList<HtmlInlineBreakProgress> BreakProgress { get; }
    internal bool SupportsContinuationReflow { get; }
    internal HtmlRenderFlowBlock? InterruptedFlow { get; }
    internal IReadOnlyList<HtmlFloatExclusion> FloatExclusions { get; }
    // Float fragmentation belongs to the inline chunk that paints the float.
    // Carry it once through wrappers instead of rediscovering shared placements.
    internal IReadOnlyList<HtmlRenderForcedBreak> ForcedBreaks { get; }
    internal IReadOnlyList<HtmlRenderLineBreakGroup> LineBreakGroups { get; }
    // Block content flattened into a table cell retains its authored keeps.
    internal IReadOnlyList<HtmlRenderAvoidBreakRange> AvoidBreakRanges { get; }
    internal IReadOnlyList<HtmlRenderContinuationGroup> ContinuationGroups { get; }
    internal IReadOnlyList<HtmlRenderTrailingGroup> TrailingGroups { get; }
}
