namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Chooses a source fragment and the repeated table groups that reserve space on its page.
    /// </summary>
    /// <param name="block">The retained source flow.</param>
    /// <param name="start">The source offset already consumed.</param>
    /// <param name="rawAvailable">The remaining destination height before repeat reservations.</param>
    /// <param name="fullPageHeight">The usable body height of an empty page.</param>
    /// <param name="canForceFragment">Allows repeat suppression and forced fragmentation when the caller cannot move to a page with more room.</param>
    /// <param name="consumedBoundaries">Optional cursor retaining distinct authored boundaries at the same offset.</param>
    /// <param name="includeProjectedBoundaries">Includes physical boundaries whose source projection already accounts for page-entry spacing.</param>
    /// <returns>Null when no safe progress fits and the caller should start another page.</returns>
    /// <remarks>The returned end covers body source only; the caller consumes the original footer after painting its final repetition.</remarks>
    private PagedBlockFragment? PlanPagedFragment(
        HtmlRenderFlowBlock block,
        double start,
        double rawAvailable,
        double fullPageHeight,
        bool canForceFragment,
        ISet<object>? consumedBoundaries = null,
        bool includeProjectedBoundaries = false) {
        CheckCancellation();
        ChargeLayoutOperations((long)block.ContinuationGroups.Count + block.TrailingGroups.Count,
            "page fragment repeat-group planning");
        HtmlRenderContinuationGroup? continuationGroup = block.ContinuationGroups.FirstOrDefault(group => group.AppliesAt(start));
        bool repeatContinuation = start > 0.0001D && continuationGroup != null && continuationGroup.Visuals.Count > 0 && continuationGroup.Height > 0D;
        double continuationHeight = repeatContinuation ? continuationGroup!.Height : 0D;
        HtmlRenderTrailingGroup? trailingGroup = ResolveTrailingGroup(block, start, Math.Max(0D, rawAvailable - continuationHeight), fullPageHeight, out double fragmentLimit);
        bool repeatTrailing = trailingGroup != null && trailingGroup.Visuals.Count > 0 && trailingGroup.Height > 0D;
        double trailingHeight = repeatTrailing ? trailingGroup!.Height : 0D;
        double available = rawAvailable - continuationHeight - trailingHeight;
        bool suppressedHeader = false;
        bool suppressedFooter = false;
        bool forced = false;
        if (consumedBoundaries != null)
            ChargeLayoutOperations(block.ForcedBreaks.Count, "page fragment authored-boundary cursor");
        bool hasForcedBreak = TryGetNextForcedBreak(block.ForcedBreaks, start, out HtmlRenderForcedBreak? forcedBreak,
            consumedBoundaries, includeProjectedBoundaries);
        // A separate empty source boundary advances the physical page without
        // consuming table rows or repeating their header/footer on a blank page.
        if (includeProjectedBoundaries && hasForcedBreak && forcedBreak!.Offset <= start + 0.0001D)
            return new PagedBlockFragment(null, null, start, rawAvailable, false, false, false, forcedBreak);
        bool forcedBreakFits = hasForcedBreak
            && forcedBreak!.Offset <= fragmentLimit + 0.0001D
            && forcedBreak.Offset <= start + available + 0.0001D;
        double fragmentEnd = forcedBreakFits
            ? forcedBreak!.Offset
            : available > 0.0001D
                ? FindFragmentEnd(block, start, available, fragmentLimit, fullPageHeight, repeatTrailing ? trailingGroup : null)
                : start;
        if (fragmentEnd <= start + 0.0001D && !forcedBreakFits) {
            if (!canForceFragment) return null;

            bool originalContinuation = repeatContinuation;
            bool originalTrailing = repeatTrailing;
            bool foundFallback = false;
            if (originalContinuation) {
                double candidateAvailable = rawAvailable - trailingHeight;
                double candidateEnd = candidateAvailable > 0.0001D
                    ? FindFragmentEnd(block, start, candidateAvailable, fragmentLimit, fullPageHeight, repeatTrailing ? trailingGroup : null)
                    : start;
                if (candidateEnd > start + 0.0001D) {
                    repeatContinuation = false;
                    continuationHeight = 0D;
                    available = candidateAvailable;
                    fragmentEnd = candidateEnd;
                    foundFallback = true;
                }
            }

            if (!foundFallback && originalTrailing) {
                double candidateAvailable = rawAvailable - (originalContinuation ? continuationGroup!.Height : 0D);
                double candidateEnd = candidateAvailable > 0.0001D
                    ? FindFragmentEnd(block, start, candidateAvailable, fragmentLimit, fullPageHeight)
                    : start;
                if (candidateEnd > start + 0.0001D) {
                    repeatContinuation = originalContinuation;
                    continuationHeight = repeatContinuation ? continuationGroup!.Height : 0D;
                    repeatTrailing = false;
                    trailingHeight = 0D;
                    available = candidateAvailable;
                    fragmentEnd = candidateEnd;
                    foundFallback = true;
                }
            }

            if (!foundFallback && originalContinuation && originalTrailing) {
                double candidateEnd = FindFragmentEnd(block, start, rawAvailable, fragmentLimit, fullPageHeight);
                if (candidateEnd > start + 0.0001D) {
                    repeatContinuation = false;
                    continuationHeight = 0D;
                    repeatTrailing = false;
                    trailingHeight = 0D;
                    available = rawAvailable;
                    fragmentEnd = candidateEnd;
                    foundFallback = true;
                }
            }

            if (foundFallback) {
                suppressedHeader = originalContinuation && !repeatContinuation;
                suppressedFooter = originalTrailing && !repeatTrailing;
            } else {
                repeatContinuation = false;
                continuationHeight = 0D;
                repeatTrailing = false;
                trailingHeight = 0D;
                available = Math.Max(0D, rawAvailable);
                fragmentEnd = Math.Min(fragmentLimit, start + available);
                forced = true;
            }
        }

        return new PagedBlockFragment(
            repeatContinuation ? continuationGroup : null,
            repeatTrailing ? trailingGroup : null,
            fragmentEnd,
            available,
            suppressedHeader,
            suppressedFooter,
            forced,
            forcedBreakFits ? forcedBreak : null);
    }

    /// <summary>Reports the fallback losses of a fragment after its page plan has been accepted.</summary>
    private void ReportPagedFragmentLoss(HtmlRenderFlowBlock block, PagedBlockFragment fragment) {
        if (fragment.SuppressedHeader) {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed, "A repeated table header was suppressed because it left no safe body-row break on an empty page.", HtmlDiagnosticSeverity.Warning, block.Source);
        }

        if (fragment.SuppressedFooter) {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TableFooterRepeatSuppressed, "A repeated table footer was suppressed because it left no safe body-row break on an empty page.", HtmlDiagnosticSeverity.Warning, block.Source);
        }

        if (fragment.Forced) {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ForcedFragment, "A layout block had no safe break opportunity within one page and was force-fragmented.", HtmlDiagnosticSeverity.Warning, block.Source);
        }
    }

    private sealed record PagedBlockFragment(
        HtmlRenderContinuationGroup? Header,
        HtmlRenderTrailingGroup? Footer,
        double End,
        double Available,
        bool SuppressedHeader,
        bool SuppressedFooter,
        bool Forced,
        HtmlRenderForcedBreak? AuthoredBreak) {
        internal double HeaderHeight => Header?.Height ?? 0D;
        internal double FooterHeight => Footer?.Height ?? 0D;
    }
}
