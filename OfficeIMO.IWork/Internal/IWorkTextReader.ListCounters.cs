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
    // or rejected metadata ends it. Qualified non-tiered decimal levels reset
    // every descendant whenever an ancestor advances or explicitly restarts.
    private sealed class ParagraphListCounter {
        private ulong? _identifier;
        private readonly int?[] _numbers = new int?[IWorkDecimalListProfile.MaximumLevel + 1];

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
            // Levels persist across sparse native boundaries. The restart flag
            // belongs only to the paragraph at its boundary: Pages omits a new
            // level entry when inserting a same-level continuation after it.
            if (levelBoundary?.Index == offset && levelBoundary.SecondaryValue == 1
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
                ClearDeeperLevels(level);
                return (_numbers[level], _numbers[level]);
            }
            if (raw == 0 && identifier == _identifier) {
                if (_numbers[level] is int number && number < int.MaxValue) {
                    _numbers[level] = number + 1;
                    ClearDeeperLevels(level);
                    return (_numbers[level], null);
                }
                if (level > 0 && !_numbers[level].HasValue && _numbers[level - 1].HasValue) {
                    _numbers[level] = 1;
                    ClearDeeperLevels(level);
                    return (1, null);
                }
            }
            complete = false;
            Reset();
            return (null, null);
        }

        private void ClearDeeperLevels(int level) {
            for (int child = level + 1; child < _numbers.Length; child++) _numbers[child] = null;
        }

        private void Reset() { _identifier = null; Array.Clear(_numbers, 0, _numbers.Length); }
    }
}
