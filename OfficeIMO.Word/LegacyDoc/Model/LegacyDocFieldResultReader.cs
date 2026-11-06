namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Retains the formatted, revision-aware text runs inside an editable field cache.</summary>
    internal sealed class LegacyDocFieldResultReader {
        private readonly IReadOnlyList<LegacyDocCharacterFormatRange> _ranges;
        private readonly bool _ordered;
        private readonly int[]? _prefixMaximumEnds;

        internal LegacyDocFieldResultReader(IReadOnlyList<LegacyDocCharacterFormatRange> ranges) {
            _ranges = ranges;
            _ordered = true;
            bool sortedStarts = true;
            int previousStart = int.MinValue;
            int previousEnd = int.MinValue;
            foreach (var range in ranges) {
                if (range.FileOffsetStart < previousStart) sortedStarts = false;
                if (range.FileOffsetStart < previousEnd || range.FileOffsetEnd <= range.FileOffsetStart) {
                    _ordered = false;
                }
                previousStart = range.FileOffsetStart;
                previousEnd = range.FileOffsetEnd;
            }
            // The Word97 reader sorts starts but can retain overlapping carriers.
            // Prefix maxima locate their earliest containing range without a per-character scan.
            if (!_ordered && sortedStarts) {
                _prefixMaximumEnds = new int[ranges.Count];
                int maximum = int.MinValue;
                for (int index = 0; index < ranges.Count; index++) {
                    maximum = Math.Max(maximum, ranges[index].FileOffsetEnd);
                    _prefixMaximumEnds[index] = maximum;
                }
            }
        }

        internal LegacyDocTextRun Read(
            IReadOnlyList<LegacyDocTextCharacter> characters, int start, int end,
            LegacyDocFieldKind kind, string? instruction, Func<char, char?> normalize,
            LegacyDocHyperlinkTarget hyperlink = default) {
            var text = new System.Text.StringBuilder();
            var positions = new List<int>();
            var runText = new System.Text.StringBuilder();
            var runPositions = new List<int>();
            var runs = new List<LegacyDocTextRun>();
            LegacyDocCharacterFormat firstFormat = LegacyDocCharacterFormat.Default;
            LegacyDocCharacterFormat currentFormat = LegacyDocCharacterFormat.Default;
            for (int index = start; index < end; index++) {
                var character = characters[index];
                char? value = normalize(character.Character);
                if (value == null) continue;
                var format = FormatAt(character.FileOffset);
                if (positions.Count == 0) firstFormat = format;
                if (runText.Length > 0 && !format.Equals(currentFormat)) FlushRun();
                currentFormat = format;
                text.Append(value.Value);
                positions.Add(character.CharacterPosition);
                runText.Append(value.Value);
                runPositions.Add(character.CharacterPosition);
            }
            FlushRun();
            bool simplePageNumber = kind == LegacyDocFieldKind.Page
                && runs.Count <= 1 && !firstFormat.Revision.HasValue;
            return LegacyDocTextRunFactory.CreateFieldRun(simplePageNumber ? string.Empty : text.ToString(),
                kind, instruction, firstFormat, positions, hyperlink, simplePageNumber ? null : runs);

            void FlushRun() {
                if (runText.Length == 0) return;
                runs.Add(LegacyDocTextRunFactory.CreateFieldRun(runText.ToString(), LegacyDocFieldKind.None,
                    currentFormat, runPositions));
                runText.Clear();
                runPositions.Clear();
            }
        }

        internal static char? NormalizeStoryCharacter(char character) =>
            char.IsControl(character) && !LegacyDocSpecialCharacters.IsSupportedInlineControl(character)
                ? null : character;

        private LegacyDocCharacterFormat FormatAt(int offset) {
            if (_ordered) {
                int index = LegacyDocDocument.FindOrderedRange(_ranges, offset,
                    static range => range.FileOffsetStart, static range => range.FileOffsetEnd);
                return index < 0 ? LegacyDocCharacterFormat.Default : _ranges[index].Format;
            }
            if (_prefixMaximumEnds != null) {
                int first = 0;
                int last = _prefixMaximumEnds.Length;
                while (first < last) {
                    int middle = first + (last - first) / 2;
                    if (_prefixMaximumEnds[middle] <= offset) first = middle + 1;
                    else last = middle;
                }
                return first < _ranges.Count && _ranges[first].Contains(offset)
                    ? _ranges[first].Format : LegacyDocCharacterFormat.Default;
            }
            foreach (var range in _ranges) {
                if (range.Contains(offset)) return range.Format;
            }
            return LegacyDocCharacterFormat.Default;
        }
    }
}
