using System.Text.RegularExpressions;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static bool TryFindLookupPosition(FormulaArgumentValue lookupValue,
            IReadOnlyList<FormulaArgumentValue> values, int matchMode, int searchMode,
            out int position, out FormulaArgumentValue failure, bool sortedApproximate = false) {
            position = 0;
            failure = default;
            bool ordered = Math.Abs(searchMode) == 2 || matchMode == -1 || matchMode == 1;
            if (ordered && LookupValueKind(lookupValue) == 1 && !TryGetArraySortKeyKind(lookupValue, out _)) {
                failure = FormulaArgumentValue.UnresolvedFormula();
                return false;
            }
            if (Math.Abs(searchMode) == 2) {
                if (matchMode == 2) { failure = FormulaArgumentValue.UnresolvedFormula(); return false; }
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
                if (candidate.IsError || candidate.IsUnresolvedFormula) { failure = candidate; return false; }
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
            int lower = 0, upper = values.Count - 1;
            while (lower <= upper) {
                int middle = lower + (upper - lower) / 2;
                FormulaArgumentValue candidate = values[middle];
                if (candidate.IsError || candidate.IsUnresolvedFormula) { failure = candidate; return false; }
                if (LookupValuesEqual(candidate, lookupValue)) { position = middle + 1; return true; }
                if (LookupValueKind(candidate) == 1 && !TryGetArraySortKeyKind(candidate, out _)) {
                    failure = FormulaArgumentValue.UnresolvedFormula();
                    return false;
                }
                int comparison = CompareLookupValues(candidate, lookupValue);
                if (matchMode != 0 && LookupValueKind(candidate) == LookupValueKind(lookupValue)
                    && (matchMode < 0 ? comparison < 0 : comparison > 0)) position = middle + 1;
                if ((searchMode == 2 ? comparison : -comparison) < 0) lower = middle + 1;
                else upper = middle - 1;
            }
            return position != 0;
        }

        private static int LookupValueKind(FormulaArgumentValue value) =>
            value.IsBoolean ? 2 : value.Number.HasValue || !value.HasValue ? 0 : 1;

        private static int CompareLookupValues(FormulaArgumentValue left, FormulaArgumentValue right) {
            int kind = LookupValueKind(left), otherKind = LookupValueKind(right);
            if (kind != otherKind) return kind.CompareTo(otherKind);
            return kind == 1
                ? StringComparer.OrdinalIgnoreCase.Compare(left.Text ?? string.Empty, right.Text ?? string.Empty)
                : (left.Number ?? 0d).CompareTo(right.Number ?? 0d);
        }
    }
}
