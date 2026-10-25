using System.Collections.ObjectModel;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Commits generated inline lines through the caller's bounded admission.
    /// Token-breaking helpers share ICollection, so the guard also stops a
    /// single overlong token instead of first allocating all of its lines.
    /// Ordinary body layout has no additional limit or cluster policy here.
    /// </summary>
    private sealed class InlineLineCollection : Collection<InlineLine> {
        private readonly int _maximumLines;
        internal InlineLineCollection(List<InlineLine> lines, int maximumLines, bool retainPreparedTokenClusters)
            : base(lines) {
            _maximumLines = maximumLines;
            RetainPreparedTokenClusters = retainPreparedTokenClusters;
        }
        internal bool RetainPreparedTokenClusters { get; }
        protected override void InsertItem(int index, InlineLine item) {
            if (Count >= _maximumLines) throw new InlineLineLimitException();
            base.InsertItem(index, item);
        }
    }

    // This private admission signal unwinds token helpers to the line-layout
    // owner, which retains admitted lines and emits the caller's typed loss.
    private sealed class InlineLineLimitException : Exception { }
}
