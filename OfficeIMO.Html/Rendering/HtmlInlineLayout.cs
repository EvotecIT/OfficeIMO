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
        double? pagedPaintExtent = null) {
        Visuals = new List<HtmlRenderVisual>(visuals);
        Height = height;
        NormalFlowHeight = normalFlowHeight ?? height;
        PagedPaintExtent = Math.Max(height, pagedPaintExtent ?? height);
        BreakOffsets = new List<double>(breakOffsets ?? Array.Empty<double>()).AsReadOnly();
        LineBreakOffsets = new List<double>(lineBreakOffsets ?? BreakOffsets).AsReadOnly();
        RunningStringAssignments = new List<HtmlCssRunningStringAssignment>(
            runningStringAssignments ?? Array.Empty<HtmlCssRunningStringAssignment>()).AsReadOnly();
        BreakProgress = new List<HtmlInlineBreakProgress>(breakProgress ?? Array.Empty<HtmlInlineBreakProgress>()).AsReadOnly();
        ForcedBreaks = forcedBreaks == null ? Array.Empty<HtmlRenderForcedBreak>() : new List<HtmlRenderForcedBreak>(forcedBreaks).AsReadOnly();
        LineBreakGroups = lineBreakGroups == null ? Array.Empty<HtmlRenderLineBreakGroup>() : new List<HtmlRenderLineBreakGroup>(lineBreakGroups).AsReadOnly();
        ContinuationGroups = continuationGroups == null ? Array.Empty<HtmlRenderContinuationGroup>() : new List<HtmlRenderContinuationGroup>(continuationGroups).AsReadOnly();
        TrailingGroups = trailingGroups == null ? Array.Empty<HtmlRenderTrailingGroup>() : new List<HtmlRenderTrailingGroup>(trailingGroups).AsReadOnly();
        SupportsContinuationReflow = supportsContinuationReflow;
        InterruptedFlow = interruptedFlow;
        FloatExclusions = new List<HtmlFloatExclusion>(floatExclusions ?? Array.Empty<HtmlFloatExclusion>()).AsReadOnly();
    }

    internal IReadOnlyList<HtmlRenderVisual> Visuals { get; }
    internal double Height { get; }
    internal double NormalFlowHeight { get; }
    // Shared float contexts can paint beyond this chunk's normal-flow height.
    internal double PagedPaintExtent { get; }
    /// <summary>Page-break candidates; may exclude line ends inside a floated box.</summary>
    internal IReadOnlyList<double> BreakOffsets { get; }
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
    internal IReadOnlyList<HtmlRenderContinuationGroup> ContinuationGroups { get; }
    internal IReadOnlyList<HtmlRenderTrailingGroup> TrailingGroups { get; }
}

