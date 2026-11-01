using System.Text.RegularExpressions;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static bool TryFindLookupPosition(FormulaArgumentValue lookupValue,
            IReadOnlyList<FormulaArgumentValue> values, int matchMode, int searchMode,
            out int position, out FormulaArgumentValue failure, bool sortedApproximate = false) {
            position = 0;
            failure = default;
            if (Math.Abs(searchMode) == 2 && matchMode == 2) {
                failure = FormulaArgumentValue.Error("#VALUE!");
                return false;
            }
            bool ordered = Math.Abs(searchMode) == 2 || matchMode == -1 || matchMode == 1;
            if (ordered && LookupValueKind(lookupValue) == 1 && !TryGetArraySortKeyKind(lookupValue, out _)) {
                failure = FormulaArgumentValue.UnresolvedFormula();
                return false;
            }
            if (Math.Abs(searchMode) == 2) {
                return TryFindBinaryLookupPosition(lookupValue, values, matchMode, searchMode, out position, out failure);
            }
            Regex? wildcard = matchMode == 2 && !lookupValue.Number.HasValue && lookupValue.Text != null
                && (lookupValue.Text.IndexOf('*') >= 0 || lookupValue.Text.IndexOf('?') >= 0 || lookupValue.Text.IndexOf('~') >= 0)
                ? new Regex("\\A" + CreateFormulaWildcardPattern(lookupValue.Text) + "\\z",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Singleline, FormulaRegexTimeout)
                : null;
            int start = searchMode == -1 ? values.Count - 1 : 0;
            int end = searchMode == -1 ? -1 : values.Count;
            int step = searchMode == -1 ? -1 : 1;
            FormulaArgumentValue best = default;
            for (int index = start; index != end; index += step) {
                FormulaArgumentValue candidate = values[index];
                if (candidate.IsUnresolvedFormula) { failure = candidate; return false; }
                if (candidate.IsError) continue;
                bool exact = wildcard != null
                    ? !candidate.Number.HasValue && candidate.Text != null && wildcard.IsMatch(candidate.Text)
                    : LookupValuesEqual(candidate, lookupValue);
                if (exact) {
                    position = index + 1;
                    if (!sortedApproximate) return true;
                    best = candidate;
                    continue;
                }
                if (matchMode == 0 || matchMode == 2 || LookupValueKind(candidate) != LookupValueKind(lookupValue)) continue;
                if (LookupValueKind(candidate) == 1 && !TryGetArraySortKeyKind(candidate, out _)) {
                    failure = FormulaArgumentValue.UnresolvedFormula();
                    return false;
                }
                int comparison = CompareLookupValues(candidate, lookupValue);
                bool eligible = matchMode < 0 ? comparison < 0 : comparison > 0;
                if (!eligible) {
                    if (sortedApproximate) break;
                    continue;
                }
                int againstBest = position == 0 ? 0 : CompareLookupValues(candidate, best);
                if (position == 0 || (matchMode < 0 ? againstBest > 0 : againstBest < 0)
                    || sortedApproximate && againstBest == 0) {
                    position = index + 1;
                    best = candidate;
                }
            }
            return position != 0;
        }

        private static bool TryFindBinaryLookupPosition(FormulaArgumentValue lookupValue,
            IReadOnlyList<FormulaArgumentValue> values, int matchMode, int searchMode,
            out int position, out FormulaArgumentValue failure) {
            position = 0;
            failure = default;
            // Search in ascending value order, reversing the source index for a
            // descending vector. The lower bound selects the native duplicate
            // boundary and leaves adjacent candidates for approximate modes.
            int lower = 0, upper = values.Count;
            int previousPosition = 0, nextPosition = 0;
            FormulaArgumentValue previous = default, next = default;
            while (lower < upper) {
                int middle = lower + (upper - lower) / 2;
                int index = searchMode == 2 ? middle : values.Count - 1 - middle;
                FormulaArgumentValue candidate = values[index];
                if (candidate.IsError || candidate.IsUnresolvedFormula) { failure = candidate; return false; }
                if (LookupValueKind(candidate) == 1 && !TryGetArraySortKeyKind(candidate, out _)) {
                    failure = FormulaArgumentValue.UnresolvedFormula();
                    return false;
                }
                int comparison = CompareLookupValues(candidate, lookupValue);
                if (comparison < 0) {
                    previous = candidate;
                    previousPosition = index + 1;
                    lower = middle + 1;
                } else {
                    next = candidate;
                    nextPosition = index + 1;
                    upper = middle;
                }
            }
            if (nextPosition != 0 && LookupValuesEqual(next, lookupValue)) position = nextPosition;
            else if (matchMode < 0 && previousPosition != 0 && LookupValueKind(previous) == LookupValueKind(lookupValue)) position = previousPosition;
            else if (matchMode > 0 && nextPosition != 0 && LookupValueKind(next) == LookupValueKind(lookupValue)) position = nextPosition;
            return position != 0;
        }

        private static int LookupValueKind(FormulaArgumentValue value) =>
            !value.HasValue ? 3 : value.IsBoolean ? 2 : value.Number.HasValue ? 0 : 1;

        private static int CompareLookupValues(FormulaArgumentValue left, FormulaArgumentValue right) {
            int kind = LookupValueKind(left), otherKind = LookupValueKind(right);
            if (kind != otherKind) return kind.CompareTo(otherKind);
            return kind == 1
                ? StringComparer.OrdinalIgnoreCase.Compare(left.Text ?? string.Empty, right.Text ?? string.Empty)
                : (left.Number ?? 0d).CompareTo(right.Number ?? 0d);
        }
    }
}
