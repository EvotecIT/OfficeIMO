namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock QualifySharedTableRepetition(HtmlRenderFlowBlock block) {
        if (block.ContinuationGroups.Count == 0 && block.TrailingGroups.Count == 0) return block;
        var headers = block.ContinuationGroups.Select(group =>
            (Group: group, Table: RepeatedTable(group.Visuals))).ToArray();
        var owners = headers.Select(header => (Owner: header.Table?.StructureElementKey, Start: header.Group.StartsAfter, End: header.Group.EndsAt))
            .Concat(block.TrailingGroups.Select(group =>
                (Owner: RepeatedTable(group.Visuals)?.StructureElementKey, Start: group.StartsAt, End: group.SourceEndsAt))).ToArray();
        var suppressed = new List<SuppressedTableRepeat>();
        var continuations = new List<HtmlRenderContinuationGroup>();
        foreach (var header in headers) {
            bool parallel = false;
            foreach (var other in owners) {
                ChargeLayoutOperation("parallel table header ownership");
                if (header.Table?.StructureElementKey != other.Owner
                    && header.Group.StartsAfter < other.End - 0.0001D
                    && other.Start < header.Group.EndsAt - 0.0001D) {
                    parallel = true;
                    break;
                }
            }
            if (parallel) suppressed.Add(new SuppressedTableRepeat(header.Table?.StructureElementKey, header.Table?.Source, true,
                header.Group.StartsAfter, header.Group.EndsAt));
            else continuations.Add(header.Group);
        }

        var trailing = new List<HtmlRenderTrailingGroup>();
        foreach (HtmlRenderTrailingGroup group in block.TrailingGroups) {
            HtmlRenderSemanticGroup? table = RepeatedTable(group.Visuals);
            string? owner = table?.StructureElementKey;
            // Consuming a repeated footer advances the shared source offset past
            // its original footer. That interval must contain no other content.
            // Assess the complete scene so later anonymous chunks and ordinary
            // paragraphs participate alongside other floats.
            bool parallel = owner != null && HasForeignContentInTableFooterInterval(
                block.Visuals, owner, group.ContentEndsAt, group.SourceEndsAt);
            if (parallel) suppressed.Add(new SuppressedTableRepeat(owner, table?.Source, false, group.StartsAt, group.SourceEndsAt));
            else trailing.Add(group);
        }
        if (suppressed.Count == 0) return block;
        HtmlRenderFlowBlock qualified = block.WithVisuals(block.Visuals, continuationGroups: continuations, trailingGroups: trailing);
        _suppressedParallelTableRepeats[qualified] = suppressed;
        return qualified;
    }

    private static HtmlRenderSemanticGroup? RepeatedTable(IEnumerable<HtmlRenderVisual> visuals) =>
        EnumeratePageFloatVisuals(visuals).OfType<HtmlRenderSemanticGroup>()
            .FirstOrDefault(group => group.Role == HtmlRenderSemanticGroupRole.Table);

    private bool HasForeignContentInTableFooterInterval(IEnumerable<HtmlRenderVisual> visuals,
        string owner, double start, double end, double verticalTranslation = 0D) {
        foreach (HtmlRenderVisual visual in visuals) {
            ChargeLayoutOperation("parallel table footer ownership");
            if (visual is HtmlRenderSemanticGroup table
                && table.Role == HtmlRenderSemanticGroupRole.Table && table.StructureElementKey == owner) continue;
            IReadOnlyList<HtmlRenderVisual>? children = visual switch {
                HtmlRenderClipGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals,
                HtmlRenderLogicalTextGroup group => group.Visuals,
                HtmlRenderPathClipGroup group => group.Visuals,
                HtmlRenderLayoutRegion group => group.Visuals,
                HtmlRenderSemanticGroup group => group.Visuals,
                _ => null
            };
            if (children != null) {
                double childTranslation = verticalTranslation;
                if (visual is HtmlRenderEffectGroup effect && TryGetVerticalPaintTranslation(effect.Transform, out double translation)) {
                    childTranslation += translation;
                }
                if (HasForeignContentInTableFooterInterval(children, owner, start, end, childTranslation)) return true;
            } else if (visual is HtmlRenderText or HtmlRenderImage or HtmlRenderDrawing or HtmlRenderFormField
                or HtmlRenderShape { IsAtomicReplacedPlaceholder: true }) {
                double top = visual.LayoutY + verticalTranslation;
                if (top < end - 0.0001D && top + visual.LayoutHeight > start + 0.0001D) return true;
            }
        }
        return false;
    }

    private void ReportParallelTableRepeatSuppressed(HtmlRenderFlowBlock block, string? owner, string? source, bool header) {
        string code = header ? HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed : HtmlRenderDiagnosticCodes.TableFooterRepeatSuppressed;
        string detail = "parallel-table-repetition;" + (owner ?? block.Source);
        if (!_reportedParallelTableRepeats.Add(code + ";" + detail)) return;
        _diagnostics.Add(ComponentName, code,
            header ? "Parallel table headers retain their original source occurrence because independent repetition offsets are not qualified."
                : "A table footer retains its original source occurrence because repeating it would skip neighboring content.",
            HtmlDiagnosticSeverity.Warning, source ?? block.Source, detail, OfficeConversionLossKind.Approximation);
    }

    private void ReportSharedTableRepetitionLoss(HtmlRenderFlowBlock block, double fragmentEnd) {
        if (!_suppressedParallelTableRepeats.TryGetValue(block, out IReadOnlyList<SuppressedTableRepeat>? groups)) return;
        foreach (SuppressedTableRepeat group in groups) {
            // Eligibility filtering is conservative. Report loss only when an
            // actual page boundary fragments the rejected group's source span.
            if (fragmentEnd > group.Start + 0.0001D && fragmentEnd < group.End - 0.0001D) {
                ReportParallelTableRepeatSuppressed(block, group.Owner, group.Source, group.Header);
            }
        }
    }

    private sealed record SuppressedTableRepeat(string? Owner, string? Source, bool Header, double Start, double End);
}
