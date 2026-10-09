namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    private static void ValidateListStartBoundaries(IReadOnlyList<AttributeBoundary> starts,
        string text, IWorkArchiveRecord owner, IWorkSourceReferenceIssueCollector references,
        ref bool complete) {
        if (starts.Count == 0) return;
        foreach (AttributeBoundary boundary in starts) {
            if (IsParagraphStart(text, boundary.Index)
                || boundary.Index == text.Length && boundary.Identifier == 0) continue;
            complete = false;
            references.Declarations.Record(owner, CounterPath(boundary) + "/1", 1,
                IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
        }
    }

    private static bool IsParagraphStart(string text, int offset) =>
        offset == 0 && text.Length > 0 || offset > 0 && offset <= text.Length
        && BreakKind(text[offset - 1]) != IWorkParagraphBreakKind.None
        && !(text[offset - 1] == '\r' && offset < text.Length && text[offset] == '\n');

    private static string CounterPath(AttributeBoundary boundary) => "14/1["
        + boundary.Position.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";

    // A counter belongs to one contiguous list in one text storage. A positive start
    // is qualified only on its own paragraph; zero continues an already-qualified
    // counter. Continuation across non-list text and other levels remains unassessed.
    private sealed class ParagraphListCounter {
        private ulong? _identifier;
        private int? _number;

        internal (int? Number, int? Start) Resolve(IReadOnlyList<AttributeBoundary> starts,
            int offset, ulong? identifier, IWorkListMarkerKind kind, bool qualifiedProfile,
            IWorkArchiveRecord owner, IWorkSourceReferenceIssueCollector references,
            ref bool complete) {
            int upper = UpperBound(starts, offset);
            AttributeBoundary? boundary = upper == 0 ? null : starts[upper - 1];
            ulong? raw = boundary?.HasObject == true ? boundary.Identifier : null;
            bool explicitStart = boundary?.Index == offset && raw > 0;
            if (kind != IWorkListMarkerKind.Number) {
                // Apple also writes a default start for bullets and non-list text.
                // It has no numbering meaning in those profiles; the table shape
                // and paragraph boundaries have already been validated separately.
                Reset();
                return (null, null);
            }
            if (!qualifiedProfile || !raw.HasValue) {
                complete = false;
                Reset();
                return (null, null);
            }
            if (explicitStart) {
                _identifier = identifier;
                _number = (int)raw!.Value;
                return (_number, _number);
            }
            if (raw == 0 && identifier == _identifier && _number.HasValue && _number < int.MaxValue) {
                _number++;
                return (_number, null);
            }
            complete = false;
            Reset();
            return (null, null);
        }

        private void Reset() { _identifier = null; _number = null; }
    }
}
