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
        double continuationStartsAfter = 0D) {
        Width = width;
        Height = height;
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
    }

    internal double Width { get; }
    internal double Height { get; }
    internal IReadOnlyList<HtmlRenderVisual> Visuals { get; }
    internal HtmlPageBreakTarget BreakBefore { get; }
    internal HtmlPageBreakTarget BreakAfter { get; }
    internal bool AvoidBreakInside { get; }
    internal string Source { get; }
    internal IReadOnlyList<double> BreakOffsets { get; }
    internal IReadOnlyList<HtmlRenderLineBreakGroup> LineBreakGroups { get; }
    internal IReadOnlyList<HtmlRenderContinuationGroup> ContinuationGroups { get; }
    internal IReadOnlyList<HtmlRenderTrailingGroup> TrailingGroups { get; }
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
    internal HtmlInlineRun(string text, HtmlRenderBoxStyle style, string? linkUri, string source) {
        Text = text;
        Style = style;
        LinkUri = linkUri;
        Source = source;
    }

    internal string Text { get; }
    internal HtmlRenderBoxStyle Style { get; }
    internal string? LinkUri { get; }
    internal string Source { get; }
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
