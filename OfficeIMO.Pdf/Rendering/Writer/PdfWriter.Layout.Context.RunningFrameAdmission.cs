namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        // This state belongs to one unplaced block or paragraph fragment. Clear
        // it after painting content; a later fragment has its own admission.
        private sealed class RunningFrameAdmissionAttempt {
            internal readonly HashSet<RunningFrameAdmissionProfile> FailedStaticFrames = new();
        }

        private readonly record struct RunningFrameAdmissionProfile(
            PdfCreateOptions Source, PdfRunningContent? Header, PdfRunningContent? Footer,
            PdfRunningContentVariant Variant, bool Mirrored, int Column,
            bool PartialPhysicalFrame, double X, double Width, double Capacity,
            double FrameHeight, double ContinuationHeight, double ContainerReservation,
            double MeasurementTopPadding, double MeasurementBottomInset, double RequiredHeight);

        /// <summary>
        /// Visits another actual running-story frame before a strict fit failure.
        /// The caller owns its normal transition and remeasures after it returns.
        /// No story or deferred-content factory is evaluated by this method.
        /// </summary>
        private bool TryAdvanceForRunningContent(ref RunningFrameAdmissionAttempt? attempt,
            double requiredHeight, Action<bool> advance, double? frameCapacity = null,
            bool contextualBody = false) {
            cancellationToken.ThrowIfCancellationRequested();
            if (isRunningContent || columnBalanceMeasurementScope != null || currentPage == null)
                return false;
            double capacity = frameCapacity ?? GetMaximumBlockContinuationHeight();
            if (requiredHeight <= capacity + .001D) return false;

            PdfCreateOptions source = currentPageBaseOptions;
            PdfRunningContentContext context = currentPage.RunningContentContext ?? GetRunningPageContext();
            bool hasRunningDefinitions = false;
            bool usesPageContext = contextualBody;
            void Include(PdfRunningContent? story, bool mayRecur = true) {
                if (story == null) return;
                hasRunningDefinitions = true;
                usesPageContext |= mayRecur && story.UsesPageContext;
            }
            Include(source.HeaderContent);
            Include(source.FooterContent);
            if (source.DifferentOddAndEvenPagesHeaderFooter) {
                Include(source.EvenPageHeaderContent);
                Include(source.EvenPageFooterContent);
            }
            if (context.SectionPageNumber == 1 && source.DifferentFirstPageHeaderFooter) {
                // This page's first-story factory has already run. It cannot
                // recur during admission within this section, so only future
                // ordinary/even factories need a finite contextual search.
                Include(source.FirstPageHeaderContent, mayRecur: false);
                Include(source.FirstPageFooterContent, mayRecur: false);
            }
            if (!hasRunningDefinitions) return false;

            if (usesPageContext) {
                // An arbitrary factory may repeat a short height for several
                // pages before changing. Only a caller's existing finite page
                // budget can bound this search; cancellation and StartPage keep
                // enforcing the normal limits. Unbounded unknown futures retain
                // the existing conservative fit failure.
                if (!maximumGeneratedPages.HasValue && !source.MaxGeneratedPages.HasValue)
                    return false;
            } else {
                attempt ??= new RunningFrameAdmissionAttempt();
                int variantPage = source.GetHeaderFooterVariantPageNumber(context.SectionPageNumber, context.PageNumber);
                PdfRunningContentVariant variant = variantPage == 1 && source.DifferentFirstPageHeaderFooter
                    ? PdfRunningContentVariant.FirstPage
                    : source.DifferentOddAndEvenPagesHeaderFooter && variantPage % 2 == 0
                        ? PdfRunningContentVariant.EvenPages : PdfRunningContentVariant.Default;
                var profile = new RunningFrameAdmissionProfile(source,
                    source.GetRunningContentForPage(variantPage, isHeader: true),
                    source.GetRunningContentForPage(variantPage, isHeader: false), variant,
                    source.MirrorMargins && currentVisiblePageNumber % 2 == 0,
                    activeColumnFlow?.ColumnIndex ?? -1, IsPartialPhysicalColumnFrame(),
                    currentOpts.MarginLeft, width, capacity, GetFullPageContentHeight(),
                    GetMeasuredContinuationFrameHeight(), GetRunningBodyContinuationReservation(),
                    containerMeasurementTopPadding, containerMeasurementBottomInset, requiredHeight);
                // First/even/default and mirror/column geometry are separate
                // opportunities even when their story objects or heights match.
                // Repeating the complete static frame closes only this attempt.
                if (!attempt.FailedStaticFrames.Add(profile)) return false;
            }

            // A spacer-only story need not dirty its page. Retain an empty
            // admission frame so physical/section numbering actually advances.
            // Ordinary column and container transition rules still own the move.
            advance(!HasCurrentPageBodyContent());
            return true;
        }
    }
}
