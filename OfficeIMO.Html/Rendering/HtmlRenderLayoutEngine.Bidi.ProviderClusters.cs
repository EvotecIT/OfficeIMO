using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // Bidi divides an already measured source segment. Its directional groups
    // therefore own portions of those exact font pens, not proportional shares
    // of nominal scalar widths or a fresh shaping context.
    private bool TryPartitionPreparedBidiWidths(InlineSegment segment,
        IReadOnlyList<string> paintElements, IReadOnlyList<string> logicalElements,
        out IReadOnlyList<double> widths) {
        widths = Array.Empty<double>();
        if (segment.Run.PreserveContextualFallback || !HasProviderMeasurementContext(segment)
            || paintElements.Count != logicalElements.Count || logicalElements.Count == 0
            || OfficeTextElements.ContainsBidiControl(segment.LogicalText)
            || segment.Run.Style.UnicodeBidi.IndexOf("override", StringComparison.Ordinal) >= 0
            || segment.Run.Style.LetterSpacing != 0D || segment.Run.Style.WordSpacing != 0D) return false;
        OfficePreparedTextShaping? context = segment.PreparedShaping;
        if (context == null && !TryPrepareProviderText(segment.LogicalText, segment.Run.Style,
                OfficeTextElements.ResolveBaseDirection(segment.LogicalText), segment.Width,
                resolvedDirectionalRun: false, out context)) return false;
        if (context == null || !string.Equals(context.Text, segment.LogicalText, StringComparison.Ordinal)
            || !HasNonnegativePreparedAdvance(context)
            || !SamePreparedAdvance(context.Measure(GetEffectiveTextFont(segment.Run.Style).Size), segment.Width)) return false;

        var elementStarts = new int[logicalElements.Count + 1];
        for (int index = 0; index < logicalElements.Count; index++) {
            CheckCancellation();
            elementStarts[index + 1] = elementStarts[index] + logicalElements[index].Length;
        }
        var clusterStarts = new SortedSet<int>();
        for (int index = 0; index < context.Result.Glyphs.Count; index++) {
            CheckCancellation();
            clusterStarts.Add(context.Result.GetLogicalClusterStart(index));
        }
        ChargeLayoutOperations(1L + (context.Result.Glyphs.Count + logicalElements.Count + 31L) / 32L,
            "bidi provider cluster partition");
        int[] clusters = clusterStarts.Concat(new[] { context.Text.Length }).ToArray();
        var measured = new double[logicalElements.Count];
        double total = 0D;
        double fontSize = GetEffectiveTextFont(segment.Run.Style).Size;
        for (int index = 0; index + 1 < clusters.Length; index++) {
            CheckCancellation();
            int owner = Array.BinarySearch(elementStarts, clusters[index]);
            if (owner < 0 || owner >= measured.Length || Array.BinarySearch(elementStarts, clusters[index + 1]) < 0
                || !context.TrySlice(clusters[index], clusters[index + 1] - clusters[index], out OfficePreparedTextShaping? slice)
                || slice == null || !HasNonnegativePreparedAdvance(slice)) return false;
            double advance = slice.Measure(fontSize);
            if (advance < 0D || double.IsNaN(advance) || double.IsInfinity(advance)) return false;
            // A ligature may cover several source graphemes. Its complete advance
            // belongs to the first; continuation elements honestly advance zero.
            measured[owner] = advance;
            total += advance;
        }
        if (!SamePreparedAdvance(total, segment.Width)) return false;
        ChargeLayoutOperations(1L + (context.Text.Length + 31L) / 32L, "bidi provider context direction");
        int[] scalars = OfficeUnicodeBidi.ToScalars(context.Text, out int[] scalarStarts);
        OfficeUnicodeBidi.Paragraph paragraph = OfficeUnicodeBidi.Resolve(scalars,
            context.Result.Direction == OfficeTextDirection.RightToLeft ? 1 : 0, _cancellationToken);
        byte[] levels = paragraph.GetLineLevels(0, scalars.Length);
        var directions = new bool[logicalElements.Count];
        int scalarIndex = 0;
        for (int index = 0; index < directions.Length; index++) {
            CheckCancellation();
            while (scalarIndex + 1 < scalarStarts.Length && scalarStarts[scalarIndex + 1] <= elementStarts[index]) scalarIndex++;
            directions[index] = (levels[scalarIndex] & 1) != 0;
        }
        segment.RetainPreparedShaping(context, elementStarts, directions);
        widths = measured;
        return true;
    }

    private OfficePreparedTextShaping? SlicePreparedBidiGroup(InlineSegment source,
        IReadOnlyList<InlineBidiElement> logicalElements, string logicalText) {
        if (source.PreparedShaping == null || source.PreparedElementOffsets == null
            || source.PreparedElementDirections == null || logicalElements.Count == 0) return null;
        // Equal advances do not prove equal glyph context. Surrounding text or
        // an outer override can change a neutral's mirroring or a slice's order.
        foreach (InlineBidiElement element in logicalElements) {
            CheckCancellation();
            if (source.PreparedElementDirections[element.SourceElementIndex] != element.RightToLeft) return null;
        }
        // Reuse the operation-local prefix map built once during partitioning.
        // Alternating directions must not rescan the whole source per group.
        int offset = source.PreparedElementOffsets[logicalElements[0].SourceElementIndex];
        ChargeLayoutOperations(1L + (logicalText.Length + 31L) / 32L, "bidi provider glyph slice");
        return source.PreparedShaping.TrySlice(offset, logicalText.Length, out OfficePreparedTextShaping? slice)
            && slice != null && string.Equals(slice.Text, logicalText, StringComparison.Ordinal) ? slice : null;
    }

    private static bool SamePreparedAdvance(double left, double right) =>
        Math.Abs(left - right) <= Math.Max(1D, Math.Abs(right)) * 1e-9D;

    // These are the canonical programs whose signed pen convention is known to
    // the prepared owner. Unknown or negative runs retain the existing fallback.
    private static bool HasNonnegativePreparedAdvance(OfficePreparedTextShaping context) => context.Program switch {
        OfficeTrueTypeFont trueType => context.WithMetricFont(trueType)?.HasNegativeAdvance == false,
        OfficeOpenTypeCffFont cff => cff.GetShapedTextAdvance(context.Text, context.Result) >= 0L,
        _ => false
    };
}
