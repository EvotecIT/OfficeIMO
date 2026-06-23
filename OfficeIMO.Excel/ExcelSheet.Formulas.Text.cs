using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool TryEvaluateTextFunction(string function, string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (function == "FORMULATEXT") {
                if (tokens.Count != 1 || !TryGetFormulaTextArgument(tokens[0], out string formulaText)) {
                    return false;
                }

                result = new FormulaArgumentValue(null, formulaText.StartsWith("=", StringComparison.Ordinal) ? formulaText : "=" + formulaText);
                return true;
            }

            if (function == "CONCAT" || function == "CONCATENATE") {
                if (tokens.Count == 0 || !TryResolveTextArgumentValues(tokens, out var parts)) {
                    return false;
                }

                result = new FormulaArgumentValue(null, string.Concat(parts));
                return true;
            }

            if (function == "TEXTJOIN") {
                if (tokens.Count < 3
                    || !TryResolveTextArgument(tokens[0], out string delimiter)
                    || !TryResolveBooleanArgument(tokens[1], out bool ignoreEmpty)
                    || !TryResolveTextArgumentValues(tokens.Skip(2), out var parts)) {
                    return false;
                }

                if (ignoreEmpty) {
                    parts = parts.Where(part => part.Length > 0).ToList();
                }

                result = new FormulaArgumentValue(null, string.Join(delimiter, parts));
                return true;
            }

            if (function == "TEXT") {
                if (tokens.Count != 2
                    || !TryResolveFormulaArgument(tokens[0], out FormulaArgumentValue value)
                    || value.IsUnresolvedFormula
                    || !value.HasValue
                    || !TryResolveTextArgument(tokens[1], out string format)
                    || !TryFormatTextFunctionValue(value, format, out string formatted)) {
                    return false;
                }

                result = new FormulaArgumentValue(null, formatted);
                return true;
            }

            if (function == "TEXTBEFORE" || function == "TEXTAFTER") {
                return TryEvaluateTextBeforeAfterFunction(function == "TEXTBEFORE", tokens, out result);
            }

            if (function == "LEFT" || function == "RIGHT") {
                if (tokens.Count < 1
                    || tokens.Count > 2
                    || !TryResolveTextArgument(tokens[0], out string text)) {
                    return false;
                }

                int count = 1;
                if (tokens.Count == 2 && !TryGetWholeNumberArgument(tokens[1], out count)) {
                    return false;
                }

                if (count < 0) {
                    return false;
                }

                count = Math.Min(count, text.Length);
                result = new FormulaArgumentValue(null, function == "LEFT"
                    ? text.Substring(0, count)
                    : text.Substring(text.Length - count, count));
                return true;
            }

            if (function == "MID") {
                if (tokens.Count != 3
                    || !TryResolveTextArgument(tokens[0], out string text)
                    || !TryGetWholeNumberArgument(tokens[1], out int start)
                    || !TryGetWholeNumberArgument(tokens[2], out int count)
                    || start < 1
                    || count < 0) {
                    return false;
                }

                int startIndex = start - 1;
                if (startIndex >= text.Length) {
                    result = new FormulaArgumentValue(null, string.Empty);
                    return true;
                }

                count = Math.Min(count, text.Length - startIndex);
                result = new FormulaArgumentValue(null, text.Substring(startIndex, count));
                return true;
            }

            if (function == "LEN") {
                if (tokens.Count != 1 || !TryResolveTextArgument(tokens[0], out string text)) {
                    return false;
                }

                result = new FormulaArgumentValue(text.Length, text.Length.ToString(CultureInfo.InvariantCulture));
                return true;
            }

            if (function == "TRIM") {
                if (tokens.Count != 1 || !TryResolveTextArgument(tokens[0], out string text)) {
                    return false;
                }

                result = new FormulaArgumentValue(null, string.Join(" ", text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)));
                return true;
            }

            if (function == "UPPER" || function == "LOWER" || function == "PROPER") {
                if (tokens.Count != 1 || !TryResolveTextArgument(tokens[0], out string text)) {
                    return false;
                }

                string transformed = function == "UPPER"
                    ? text.ToUpperInvariant()
                    : function == "LOWER"
                        ? text.ToLowerInvariant()
                        : ToProperCase(text);
                result = new FormulaArgumentValue(null, transformed);
                return true;
            }

            if (function == "SUBSTITUTE") {
                if (tokens.Count < 3
                    || tokens.Count > 4
                    || !TryResolveTextArgument(tokens[0], out string text)
                    || !TryResolveTextArgument(tokens[1], out string oldText)
                    || !TryResolveTextArgument(tokens[2], out string newText)) {
                    return false;
                }

                if (tokens.Count == 4) {
                    if (!TryGetWholeNumberArgument(tokens[3], out int occurrence) || occurrence < 1) {
                        return false;
                    }

                    result = new FormulaArgumentValue(null, SubstituteTextOccurrence(text, oldText, newText, occurrence));
                    return true;
                }

                result = new FormulaArgumentValue(null, oldText.Length == 0 ? text : text.Replace(oldText, newText));
                return true;
            }

            if (function == "FIND" || function == "SEARCH") {
                if (tokens.Count < 2
                    || tokens.Count > 3
                    || !TryResolveTextArgument(tokens[0], out string findText)
                    || !TryResolveTextArgument(tokens[1], out string withinText)) {
                    return false;
                }

                int start = 1;
                if (tokens.Count == 3 && !TryGetWholeNumberArgument(tokens[2], out start)) {
                    return false;
                }

                if (start < 1 || start > withinText.Length + 1) {
                    return false;
                }

                StringComparison comparison = function == "SEARCH"
                    ? StringComparison.OrdinalIgnoreCase
                    : StringComparison.Ordinal;
                int foundIndex = withinText.IndexOf(findText, start - 1, comparison);
                if (foundIndex < 0) {
                    return false;
                }

                double position = foundIndex + 1;
                result = new FormulaArgumentValue(position, position.ToString(CultureInfo.InvariantCulture));
                return true;
            }

            if (function == "VALUE") {
                if (tokens.Count != 1
                    || !TryResolveTextArgument(tokens[0], out string text)
                    || !TryParseValueFunctionNumber(text, out double number)) {
                    return false;
                }

                result = new FormulaArgumentValue(number, number.ToString(CultureInfo.InvariantCulture));
                return true;
            }

            if (function == "EXACT") {
                if (tokens.Count != 2
                    || !TryResolveTextArgument(tokens[0], out string left)
                    || !TryResolveTextArgument(tokens[1], out string right)) {
                    return false;
                }

                double value = string.Equals(left, right, StringComparison.Ordinal) ? 1d : 0d;
                result = new FormulaArgumentValue(value, value.ToString(CultureInfo.InvariantCulture));
                return true;
            }

            if (function == "REPT") {
                if (tokens.Count != 2
                    || !TryResolveTextArgument(tokens[0], out string text)
                    || !TryGetWholeNumberArgument(tokens[1], out int count)
                    || count < 0) {
                    return false;
                }

                if (text.Length > 0 && count > MaxSupportedFormulaLength / text.Length) {
                    return false;
                }

                result = new FormulaArgumentValue(null, string.Concat(Enumerable.Repeat(text, count)));
                return true;
            }

            return false;
        }

        private bool TryEvaluateTextBeforeAfterFunction(bool before, IReadOnlyList<string> tokens, out FormulaArgumentValue result) {
            result = default;
            if (tokens.Count < 2
                || tokens.Count > 6
                || !TryResolveTextArgument(tokens[0], out string text)
                || !TryResolveTextArgument(tokens[1], out string delimiter)
                || delimiter.Length == 0) {
                return false;
            }

            int instance = 1;
            if (tokens.Count >= 3 && !TryGetWholeNumberArgument(tokens[2], out instance)) {
                return false;
            }

            if (instance == 0) {
                return false;
            }

            int matchMode = 0;
            if (tokens.Count >= 4 && (!TryGetWholeNumberArgument(tokens[3], out matchMode) || (matchMode != 0 && matchMode != 1))) {
                return false;
            }

            bool matchEnd = false;
            if (tokens.Count >= 5 && !TryResolveBooleanArgument(tokens[4], out matchEnd)) {
                return false;
            }

            string? ifNotFound = null;
            if (tokens.Count >= 6 && !TryResolveTextArgument(tokens[5], out ifNotFound)) {
                return false;
            }

            StringComparison comparison = matchMode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            if (!TryFindTextDelimiterOccurrence(text, delimiter, instance, comparison, out int index)) {
                if (matchEnd) {
                    result = new FormulaArgumentValue(null, before ? text : string.Empty);
                    return true;
                }

                if (ifNotFound != null) {
                    result = new FormulaArgumentValue(null, ifNotFound);
                    return true;
                }

                return false;
            }

            string extracted = before
                ? text.Substring(0, index)
                : text.Substring(index + delimiter.Length);
            result = new FormulaArgumentValue(null, extracted);
            return true;
        }

        private static bool TryFindTextDelimiterOccurrence(string text, string delimiter, int instance, StringComparison comparison, out int index) {
            if (instance > 0) {
                int searchStart = 0;
                for (int current = 1; current <= instance; current++) {
                    index = text.IndexOf(delimiter, searchStart, comparison);
                    if (index < 0) {
                        return false;
                    }

                    if (current == instance) {
                        return true;
                    }

                    searchStart = index + delimiter.Length;
                    if (searchStart > text.Length) {
                        return false;
                    }
                }
            } else {
                if (text.Length == 0) {
                    index = -1;
                    return false;
                }

                int searchStart = text.Length - 1;
                for (int current = -1; current >= instance; current--) {
                    index = text.LastIndexOf(delimiter, searchStart, comparison);
                    if (index < 0) {
                        return false;
                    }

                    if (current == instance) {
                        return true;
                    }

                    searchStart = index - 1;
                    if (searchStart < 0) {
                        return false;
                    }
                }
            }

            index = -1;
            return false;
        }

        private static string SubstituteTextOccurrence(string text, string oldText, string newText, int occurrence) {
            if (oldText.Length == 0) {
                return text;
            }

            int startIndex = 0;
            int currentOccurrence = 0;
            while (startIndex <= text.Length) {
                int index = text.IndexOf(oldText, startIndex, StringComparison.Ordinal);
                if (index < 0) {
                    return text;
                }

                currentOccurrence++;
                if (currentOccurrence == occurrence) {
                    return text.Substring(0, index) + newText + text.Substring(index + oldText.Length);
                }

                startIndex = index + oldText.Length;
            }

            return text;
        }

        private static bool TryParseValueFunctionNumber(string text, out double number) {
            string normalized = text.Trim();
            if (normalized.Length == 0) {
                number = 0d;
                return false;
            }

            normalized = normalized.Replace(",", string.Empty);
            return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out number);
        }

        private static string ToProperCase(string text) {
            var builder = new StringBuilder(text.Length);
            bool capitalizeNext = true;
            foreach (char character in text) {
                if (char.IsLetter(character)) {
                    builder.Append(capitalizeNext
                        ? char.ToUpperInvariant(character)
                        : char.ToLowerInvariant(character));
                    capitalizeNext = false;
                    continue;
                }

                builder.Append(character);
                capitalizeNext = true;
            }

            return builder.ToString();
        }

        private bool TryEvaluateIndexValue(string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (tokens.Count < 2
                || tokens.Count > 3
                || !TryResolveFormulaRangeReference(tokens[0], out ExcelSheet rangeSheet, out int r1, out int c1, out int r2, out int c2)
                || !TryGetWholeNumberArgument(tokens[1], out int rowIndex)) {
                return false;
            }

            int rowCount = r2 - r1 + 1;
            int columnCount = c2 - c1 + 1;
            int columnIndex;
            if (tokens.Count == 3) {
                if (!TryGetWholeNumberArgument(tokens[2], out columnIndex)) {
                    return false;
                }
            } else if (columnCount == 1) {
                columnIndex = 1;
            } else if (rowCount == 1) {
                columnIndex = rowIndex;
                rowIndex = 1;
            } else {
                return false;
            }

            if (rowIndex < 1 || rowIndex > rowCount || columnIndex < 1 || columnIndex > columnCount) {
                return false;
            }

            result = rangeSheet.ResolveCellArgument(r1 + rowIndex - 1, c1 + columnIndex - 1);
            return result.HasValue;
        }

        private bool TryEvaluateMatchFunction(string function, string args, out double result) {
            result = 0;
            var tokens = SplitFormulaArguments(args);
            int maxTokens = function == "XMATCH" ? 4 : 3;
            if (tokens.Count < 2
                || tokens.Count > maxTokens
                || !TryResolveFormulaArgument(tokens[0], out FormulaArgumentValue lookupValue)
                || !TryResolveFormulaRangeReference(tokens[1], out ExcelSheet rangeSheet, out int r1, out int c1, out int r2, out int c2)) {
                return false;
            }

            bool vertical = c1 == c2;
            bool horizontal = r1 == r2;
            if (!vertical && !horizontal) {
                return false;
            }

            int matchMode = function == "XMATCH" ? 0 : 1;
            if (tokens.Count >= 3 && !TryGetWholeNumberArgument(tokens[2], out matchMode)) {
                return false;
            }

            if (function == "MATCH" && matchMode != -1 && matchMode != 0 && matchMode != 1) {
                return false;
            }

            if (function == "XMATCH" && matchMode != -1 && matchMode != 0 && matchMode != 1) {
                return false;
            }

            int searchMode = 1;
            if (function == "XMATCH"
                && tokens.Count >= 4
                && (!TryGetWholeNumberArgument(tokens[3], out searchMode) || (searchMode != 1 && searchMode != -1))) {
                return false;
            }

            int lookupMode = function == "MATCH" ? -matchMode : matchMode;
            if (!TryResolveFormulaRange(tokens[1], out var lookupValues)
                || !TryFindLookupPosition(lookupValue, lookupValues, lookupMode, searchMode, out int position)) {
                return false;
            }

            result = position;
            return true;
        }

        private static bool TryFindLookupPosition(FormulaArgumentValue lookupValue, IReadOnlyList<FormulaArgumentValue> lookupValues, int matchMode, int searchMode, out int position) {
            position = 0;
            if (lookupValues.Count == 0) {
                return false;
            }

            int start = searchMode == -1 ? lookupValues.Count - 1 : 0;
            int end = searchMode == -1 ? -1 : lookupValues.Count;
            int step = searchMode == -1 ? -1 : 1;

            for (int index = start; index != end; index += step) {
                if (!FormulaValuesEqual(lookupValues[index], lookupValue)) {
                    continue;
                }

                position = index + 1;
                return true;
            }

            if (matchMode == 0 || !lookupValue.Number.HasValue) {
                return false;
            }

            double bestDelta = double.MaxValue;
            int bestPosition = 0;
            for (int index = start; index != end; index += step) {
                FormulaArgumentValue candidate = lookupValues[index];
                if (!candidate.Number.HasValue) {
                    continue;
                }

                double delta = candidate.Number.Value - lookupValue.Number.Value;
                bool eligible = matchMode < 0
                    ? delta <= 0d
                    : delta >= 0d;
                if (!eligible) {
                    continue;
                }

                double distance = Math.Abs(delta);
                if (distance < bestDelta) {
                    bestDelta = distance;
                    bestPosition = index + 1;
                }
            }

            if (bestPosition == 0) {
                return false;
            }

            position = bestPosition;
            return true;
        }

        private bool TryFormatTextFunctionValue(FormulaArgumentValue value, string format, out string formatted) {
            formatted = string.Empty;
            if (string.IsNullOrWhiteSpace(format)) {
                return false;
            }

            if (format == "@") {
                formatted = FormulaValueToText(value);
                return true;
            }

            if (LooksLikeDateTextFormat(format)) {
                if (!value.Number.HasValue || !TryGetDateTimeFromSerial(value.Number.Value, out DateTime date)) {
                    return false;
                }

                string dotNetFormat = ConvertExcelDateTextFormat(format);
                try {
                    formatted = date.ToString(dotNetFormat, CultureInfo.InvariantCulture);
                    return true;
                } catch (FormatException) {
                    formatted = string.Empty;
                    return false;
                }
            }

            if (!value.Number.HasValue || !IsSupportedTextNumericFormat(format)) {
                return false;
            }

            try {
                formatted = value.Number.Value.ToString(format, CultureInfo.InvariantCulture);
                return true;
            } catch (FormatException) {
                formatted = string.Empty;
                return false;
            }
        }

        private static bool LooksLikeDateTextFormat(string format) {
            bool inQuote = false;
            for (int index = 0; index < format.Length; index++) {
                char ch = format[index];
                if (ch == '"') {
                    inQuote = !inQuote;
                    continue;
                }

                if (!inQuote && (ch == 'y' || ch == 'Y' || ch == 'd' || ch == 'D' || ch == 'h' || ch == 'H' || ch == 's' || ch == 'S')) {
                    return true;
                }
            }

            return false;
        }

        private static string ConvertExcelDateTextFormat(string format) {
            var builder = new StringBuilder(format.Length);
            bool inQuote = false;
            for (int index = 0; index < format.Length; index++) {
                char ch = format[index];
                if (ch == '"') {
                    inQuote = !inQuote;
                    builder.Append(ch);
                    continue;
                }

                if (!inQuote && (ch == 'm' || ch == 'M')) {
                    int start = index;
                    while (index + 1 < format.Length && (format[index + 1] == 'm' || format[index + 1] == 'M')) {
                        index++;
                    }

                    int count = index - start + 1;
                    bool minute = IsMinuteTextFormatToken(format, start, index);
                    builder.Append(new string(minute ? 'm' : 'M', count));
                    continue;
                }

                if (!inQuote && (ch == 'h' || ch == 'H')) {
                    int start = index;
                    while (index + 1 < format.Length && (format[index + 1] == 'h' || format[index + 1] == 'H')) {
                        index++;
                    }

                    builder.Append(new string('H', index - start + 1));
                    continue;
                }

                builder.Append(ch);
            }

            return builder.ToString();
        }

        private static bool IsMinuteTextFormatToken(string format, int start, int end) {
            char? previous = PreviousNonSpaceCharacter(format, start - 1);
            char? next = NextNonSpaceCharacter(format, end + 1);
            return previous == ':' || next == ':';
        }

        private static char? PreviousNonSpaceCharacter(string value, int start) {
            for (int index = start; index >= 0; index--) {
                if (!char.IsWhiteSpace(value[index])) {
                    return value[index];
                }
            }

            return null;
        }

        private static char? NextNonSpaceCharacter(string value, int start) {
            for (int index = start; index < value.Length; index++) {
                if (!char.IsWhiteSpace(value[index])) {
                    return value[index];
                }
            }

            return null;
        }

        private static bool IsSupportedTextNumericFormat(string format) {
            bool inQuote = false;
            foreach (char ch in format) {
                if (ch == '"') {
                    inQuote = !inQuote;
                    continue;
                }

                if (inQuote) {
                    continue;
                }

                if (ch == '0' || ch == '#' || ch == '.' || ch == ',' || ch == '%' || ch == '$'
                    || ch == '-' || ch == '+' || ch == '(' || ch == ')' || ch == ' ') {
                    continue;
                }

                return false;
            }

            return !inQuote && format.IndexOfAny(new[] { '0', '#' }) >= 0;
        }

    }
}
