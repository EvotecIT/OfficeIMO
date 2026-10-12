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

    private static string LevelPath(AttributeBoundary boundary) => "6/1["
        + boundary.Position.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]/3";

    private static void ValidateListRestartBoundaries(IReadOnlyList<AttributeBoundary> levels,
        IReadOnlyList<AttributeBoundary> starts, string text, IWorkArchiveRecord owner,
        IWorkSourceReferenceIssueCollector references, ref bool complete) {
        foreach (AttributeBoundary boundary in levels) {
            if (boundary.SecondaryValue != 1) continue;
            int upper = UpperBound(starts, boundary.Index);
            AttributeBoundary? start = upper == 0 ? null : starts[upper - 1];
            if (IsParagraphStart(text, boundary.Index) && boundary.Index < text.Length
                && start?.Index == boundary.Index && start.HasObject && start.Identifier > 0) continue;
            complete = false;
            references.Declarations.Record(owner, LevelPath(boundary), 1,
                IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
        }
    }

    // One counter scope belongs to one selected list style in one text storage.
    // Ordinary non-list paragraphs preserve it; a different numbered/bullet style
    // or rejected metadata ends it. The observed two non-tiered decimal levels
    // reset their children whenever a parent item advances or explicitly restarts.
    private sealed class ParagraphListCounter {
        private ulong? _identifier;
        private readonly int?[] _numbers = new int?[2];

        internal (int? Number, int? Start) Resolve(IReadOnlyList<AttributeBoundary> starts,
            IReadOnlyList<AttributeBoundary> levels, int offset, ulong? identifier, int level,
            IWorkListMarkerKind kind, bool qualifiedProfile, bool qualifiedPlainProfile,
            IWorkArchiveRecord owner, IWorkSourceReferenceIssueCollector references,
            ref bool complete) {
            int upper = UpperBound(starts, offset);
            AttributeBoundary? boundary = upper == 0 ? null : starts[upper - 1];
            ulong? raw = boundary?.HasObject == true ? boundary.Identifier : null;
            bool explicitStart = boundary?.Index == offset && raw > 0;
            int levelUpper = UpperBound(levels, offset);
            AttributeBoundary? levelBoundary = levelUpper == 0 ? null : levels[levelUpper - 1];
            if (levelBoundary?.SecondaryValue == 1
                && (!explicitStart || kind != IWorkListMarkerKind.Number)) {
                complete = false;
                references.Declarations.Record(owner, LevelPath(levelBoundary), 1,
                    IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
                Reset();
                return (null, null);
            }
            if (kind != IWorkListMarkerKind.Number) {
                // Apple also writes a default start for bullets and non-list text.
                // It has no numbering meaning in those profiles; the table shape
                // and paragraph boundaries have already been validated separately.
                if (!qualifiedPlainProfile) Reset();
                return (null, null);
            }
            if (!qualifiedProfile || !raw.HasValue) {
                complete = false;
                Reset();
                return (null, null);
            }
            if (_identifier != identifier) Reset();
            if (explicitStart) {
                _identifier = identifier;
                _numbers[level] = (int)raw!.Value;
                if (level == 0) _numbers[1] = null;
                return (_numbers[level], _numbers[level]);
            }
            if (raw == 0 && identifier == _identifier) {
                if (_numbers[level] is int number && number < int.MaxValue) {
                    _numbers[level] = number + 1;
                    if (level == 0) _numbers[1] = null;
                    return (_numbers[level], null);
                }
                if (level == 1 && !_numbers[1].HasValue && _numbers[0].HasValue) {
                    _numbers[1] = 1;
                    return (1, null);
                }
            }
            complete = false;
            Reset();
            return (null, null);
        }

        private void Reset() { _identifier = null; _numbers[0] = null; _numbers[1] = null; }
    }
}
