using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Indexes decoded text intervals once and advances interaction state in
    /// document order, avoiding repeated full-list scans per paragraph and run.
    /// </summary>
    internal sealed class LegacyPptTextProjectionPlan {
        private readonly LegacyPptParagraphRun[] _paragraphRuns;
        private readonly LegacyPptCharacterRun[] _characterRuns;
        private readonly LegacyPptTextLanguageRun[] _languageRuns;
        private readonly IReadOnlyDictionary<int, LegacyPptTextField>
            _fieldsByPosition;
        private readonly IndexedInteraction[] _interactionsByStart;
        private readonly IndexedInteraction[] _interactionsByEnd;
        private readonly LegacyPptTextInteraction[] _interactionsByOriginalIndex;
        private readonly SortedSet<int> _activeClickInteractions = new();
        private readonly SortedSet<int> _activeHoverInteractions = new();
        private readonly int[] _boundaries;
        private int _interactionStartIndex;
        private int _interactionEndIndex;
        private int _lastInteractionPosition = -1;

        internal LegacyPptTextProjectionPlan(LegacyPptTextBody source) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            _paragraphRuns = source.ParagraphRuns.OrderBy(run => run.Start)
                .ToArray();
            _characterRuns = source.CharacterRuns.OrderBy(run => run.Start)
                .ToArray();
            _languageRuns = source.LanguageRuns.OrderBy(run => run.Start)
                .ToArray();
            _fieldsByPosition = source.Fields.GroupBy(field => field.Position)
                .ToDictionary(group => group.Key, group => group.First());
            _interactionsByOriginalIndex = source.Interactions.ToArray();
            IndexedInteraction[] indexed = source.Interactions.Select(
                    (interaction, index) => new IndexedInteraction(
                        interaction, index))
                .ToArray();
            _interactionsByStart = indexed.OrderBy(item => item.Start)
                .ThenBy(item => item.OriginalIndex).ToArray();
            _interactionsByEnd = indexed.OrderBy(item => item.End)
                .ThenBy(item => item.OriginalIndex).ToArray();

            var boundaries = new SortedSet<int>();
            foreach (LegacyPptCharacterRun run in _characterRuns) {
                AddBoundaries(boundaries, run.Start, run.Length,
                    source.Text.Length);
            }
            foreach (LegacyPptTextLanguageRun run in _languageRuns) {
                AddBoundaries(boundaries, run.Start, run.Length,
                    source.Text.Length);
            }
            foreach (LegacyPptTextInteraction interaction in
                     source.Interactions) {
                AddBoundaries(boundaries, interaction.Start,
                    interaction.Length, source.Text.Length);
            }
            foreach (LegacyPptTextField field in source.Fields) {
                AddBoundaries(boundaries, field.Position, length: 1,
                    source.Text.Length);
            }
            _boundaries = boundaries.ToArray();
        }

        internal LegacyPptParagraphRun? FindParagraphRun(int position) =>
            FindContaining(_paragraphRuns, position,
                run => run.Start, run => run.Length);

        internal LegacyPptCharacterRun? FindCharacterRun(int position) =>
            FindContaining(_characterRuns, position,
                run => run.Start, run => run.Length);

        internal LegacyPptTextLanguageRun? FindLanguageRun(int position) =>
            FindContaining(_languageRuns, position,
                run => run.Start, run => run.Length);

        internal LegacyPptTextField? FindField(int start, int end) =>
            end == start + 1
            && _fieldsByPosition.TryGetValue(start,
                out LegacyPptTextField? field)
                ? field
                : null;

        internal IReadOnlyList<int> GetParagraphBoundaries(int start,
            int end) {
            var result = new List<int> { start };
            int index = LowerBound(_boundaries, start + 1);
            while (index < _boundaries.Length
                   && _boundaries[index] < end) {
                result.Add(_boundaries[index++]);
            }
            if (end != start) result.Add(end);
            return result;
        }

        internal IReadOnlyList<LegacyPptInteraction> GetInteractions(
            int position) {
            if (position < _lastInteractionPosition) {
                throw new InvalidOperationException(
                    "Text projection positions must advance monotonically.");
            }
            _lastInteractionPosition = position;
            while (_interactionStartIndex < _interactionsByStart.Length
                   && _interactionsByStart[_interactionStartIndex].Start
                   <= position) {
                IndexedInteraction item =
                    _interactionsByStart[_interactionStartIndex++];
                GetActiveSet(item.Interaction.Interaction.Trigger)
                    .Add(item.OriginalIndex);
            }
            while (_interactionEndIndex < _interactionsByEnd.Length
                   && _interactionsByEnd[_interactionEndIndex].End
                   <= position) {
                IndexedInteraction item =
                    _interactionsByEnd[_interactionEndIndex++];
                GetActiveSet(item.Interaction.Interaction.Trigger)
                    .Remove(item.OriginalIndex);
            }

            var active = new List<(int Index,
                LegacyPptInteraction Interaction)>(2);
            AddFirstActive(_activeClickInteractions, active);
            AddFirstActive(_activeHoverInteractions, active);
            return active.OrderBy(item => item.Index)
                .Select(item => item.Interaction).ToArray();
        }

        private void AddFirstActive(ISet<int> active,
            ICollection<(int Index, LegacyPptInteraction Interaction)>
                result) {
            if (active.Count == 0) return;
            int index = active.Min();
            result.Add((index,
                _interactionsByOriginalIndex[index].Interaction));
        }

        private ISet<int> GetActiveSet(
            LegacyPptInteractionTrigger trigger) =>
            trigger == LegacyPptInteractionTrigger.MouseOver
                ? _activeHoverInteractions
                : _activeClickInteractions;

        private static T? FindContaining<T>(T[] runs, int position,
            Func<T, int> getStart, Func<T, int> getLength)
            where T : class {
            int low = 0;
            int high = runs.Length - 1;
            int candidate = -1;
            while (low <= high) {
                int middle = low + ((high - low) / 2);
                if (getStart(runs[middle]) <= position) {
                    candidate = middle;
                    low = middle + 1;
                } else {
                    high = middle - 1;
                }
            }
            if (candidate < 0) return null;
            T run = runs[candidate];
            long end = (long)getStart(run) + getLength(run);
            return position < end ? run : null;
        }

        private static int LowerBound(int[] values, int target) {
            int low = 0;
            int high = values.Length;
            while (low < high) {
                int middle = low + ((high - low) / 2);
                if (values[middle] < target) low = middle + 1;
                else high = middle;
            }
            return low;
        }

        private static void AddBoundaries(ISet<int> boundaries, int start,
            int length, int textLength) {
            long rawEnd = (long)start + length;
            int clippedStart = Math.Max(0, Math.Min(textLength, start));
            int clippedEnd = Math.Max(0, Math.Min(textLength,
                rawEnd > int.MaxValue ? int.MaxValue : (int)rawEnd));
            if (clippedEnd <= clippedStart) return;
            boundaries.Add(clippedStart);
            boundaries.Add(clippedEnd);
        }

        private readonly struct IndexedInteraction {
            internal IndexedInteraction(LegacyPptTextInteraction interaction,
                int originalIndex) {
                Interaction = interaction;
                OriginalIndex = originalIndex;
                Start = interaction.Start;
                End = checked(interaction.Start + interaction.Length);
            }

            internal LegacyPptTextInteraction Interaction { get; }
            internal int OriginalIndex { get; }
            internal int Start { get; }
            internal int End { get; }
        }
    }
}
