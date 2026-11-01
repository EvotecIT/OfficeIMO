using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static bool IsSupportedTextFunction(string function) => function is
            "FORMULATEXT" or "CONCAT" or "CONCATENATE" or "TEXTJOIN" or "TEXT" or "TEXTBEFORE" or "TEXTAFTER" or
            "LEFT" or "RIGHT" or "MID" or "LEN" or "TRIM" or "UPPER" or "LOWER" or "PROPER" or "SUBSTITUTE" or
            "FIND" or "SEARCH" or "VALUE" or "EXACT" or "REPT";

        private bool TryEvaluateTextFunction(string function, string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args);
            if (function == "FORMULATEXT") {
                if (tokens.Count != 1) return false;
                if (!TryGetFormulaTextArgument(tokens[0], out string formulaText)) return _formulaReferenceError.IsError;

                result = new FormulaArgumentValue(null, formulaText.StartsWith("=", StringComparison.Ordinal) ? formulaText : "=" + formulaText);
                return true;
            }

            if (function == "CONCAT" || function == "CONCATENATE") {
                if (tokens.Count == 0 || !TryResolveTextArgumentValues(tokens, out var parts, out string? error)) return false;
                result = error != null ? FormulaArgumentValue.Error(error) : CombineFormulaText(parts, string.Empty);
                return true;
            }

            if (function == "TEXTJOIN") {
                if (tokens.Count < 3 || tokens.Count > 254
                    || !TryResolveFormulaArgument(tokens[0], out FormulaArgumentValue delimiterValue)
                    || delimiterValue.IsUnresolvedFormula) return false;
                if (delimiterValue.IsError) { result = delimiterValue; return true; }
                if (!TryResolveFormulaValueArgument(tokens[1], out FormulaArgumentValue emptyPolicy)) return false;
                if (emptyPolicy.IsError) {
                    result = emptyPolicy;
                    return true;
                }
                if (!emptyPolicy.Number.HasValue
                    || !TryResolveTextArgumentValues(tokens.Skip(2), out var parts, out string? error)) return false;
                bool ignoreEmpty = Math.Abs(emptyPolicy.Number.Value) >= double.Epsilon;
                if (error != null) { result = FormulaArgumentValue.Error(error); return true; }
                if (ignoreEmpty) parts = parts.Where(part => part.Length > 0).ToList();
                result = CombineFormulaText(parts, FormulaValueToText(delimiterValue));
                return true;
            }

            if (function == "TEXT") {
                if (tokens.Count != 2
                    || !TryResolveFormulaArgument(tokens[0], out FormulaArgumentValue value)
                    || value.IsUnresolvedFormula) {
                    return false;
                }
                if (value.IsError && !_formulaReferenceError.IsError) _formulaReferenceError = value;
                if (!TryResolveTextArgument(tokens[1], out string format)) return false;
                if (_formulaReferenceError.IsError) return true;
                if (!TryFormatTextFunctionValue(NormalizeLookupResult(value), format, out string formatted)) return false;

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
                if (tokens.Count == 2 && !TryGetTextWholeNumberArgument(tokens[1], out count)) {
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
                    || !TryGetTextWholeNumberArgument(tokens[1], out int start)
                    || !TryGetTextWholeNumberArgument(tokens[2], out int count)
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
                    if (!TryGetTextWholeNumberArgument(tokens[3], out int occurrence) || occurrence < 1) {
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
                if (tokens.Count == 3 && !TryGetTextWholeNumberArgument(tokens[2], out start)) {
                    return false;
                }

                if (start < 1 || start > withinText.Length + 1) {
                    return false;
                }

                StringComparison comparison = function == "SEARCH"
                    ? StringComparison.OrdinalIgnoreCase
                    : StringComparison.Ordinal;
                int foundIndex = function == "SEARCH" && findText.IndexOfAny(new[] { '*', '?', '~' }) >= 0
                    ? FindFormulaWildcard(findText, withinText, start - 1)
                    : withinText.IndexOf(findText, start - 1, comparison);
                if (foundIndex < 0) {
                    result = FormulaArgumentValue.Error("#VALUE!");
                    return true;
                }

                double position = foundIndex + 1;
                result = new FormulaArgumentValue(position, InvariantNumberText.Get(position));
                return true;
            }

            if (function == "VALUE") {
                if (tokens.Count != 1 || !TryResolveFormulaArgument(tokens[0], out FormulaArgumentValue value)
                    || value.IsUnresolvedFormula) return false;
                if (value.IsError) { result = value; return true; }
                if (!value.HasValue) { result = new FormulaArgumentValue(0d, "0"); return true; }
                string text = FormulaValueToText(value);
                if (value.IsBoolean || text.Length == 0) { result = FormulaArgumentValue.Error("#VALUE!"); return true; }
                if (!TryParseValueFunctionNumber(text, out double number)) return false;

                result = new FormulaArgumentValue(number, InvariantNumberText.Get(number));
                return true;
            }

            if (function == "EXACT") {
                if (tokens.Count != 2
                    || !TryResolveTextArgument(tokens[0], out string left)
                    || !TryResolveTextArgument(tokens[1], out string right)) {
                    return false;
                }

                double value = string.Equals(left, right, StringComparison.Ordinal) ? 1d : 0d;
                result = new FormulaArgumentValue(value, InvariantNumberText.Get(value), isBoolean: true);
                return true;
            }

            if (function == "REPT") {
                if (tokens.Count != 2
                    || !TryResolveTextArgument(tokens[0], out string text)
                    || !TryGetTextWholeNumberArgument(tokens[1], out int count)
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

        private static FormulaArgumentValue CombineFormulaText(IReadOnlyList<string> parts, string delimiter) {
            long length = (long)delimiter.Length * Math.Max(0, parts.Count - 1);
            foreach (string part in parts) {
                length += part.Length;
                if (length > 32767) return FormulaArgumentValue.Error("#VALUE!");
            }
            return new FormulaArgumentValue(null, string.Join(delimiter, parts));
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
            if (tokens.Count >= 3 && !TryGetTextWholeNumberArgument(tokens[2], out instance)) {
                return false;
            }

            if (instance == 0) {
                return false;
            }

            int matchMode = 0;
            if (tokens.Count >= 4 && (!TryGetTextWholeNumberArgument(tokens[3], out matchMode) || (matchMode != 0 && matchMode != 1))) {
                return false;
            }

            bool matchEnd = false;
            if (tokens.Count >= 5) {
                if (!TryEvaluateCondition(tokens[4], out matchEnd, out FormulaArgumentValue matchEndError)) return false;
                if (matchEndError.IsError && !_formulaReferenceError.IsError) _formulaReferenceError = matchEndError;
            }

            if (_formulaReferenceError.IsError) return true;

            StringComparison comparison = matchMode == 1 ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            if (!TryFindTextDelimiterOccurrence(text, delimiter, instance, comparison, out int index)) {
                if (matchEnd) {
                    result = new FormulaArgumentValue(null, before ? text : string.Empty);
                    return true;
                }

                if (tokens.Count >= 6) {
                    if (!TryResolveTextArgument(tokens[5], out string ifNotFound)) return false;
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

        private bool TryEvaluateMatchValue(string function, string args, out FormulaArgumentValue result) {
            result = default;
            var tokens = SplitFormulaArguments(args, preserveEmpty: true);
            int maxTokens = function == "XMATCH" ? 4 : 3;
            if (tokens.Count < 2
                || tokens.Count > maxTokens
                || !TryResolveLookupArgument(tokens[0], out FormulaArgumentValue lookupValue)
                || lookupValue.IsUnresolvedFormula) {
                return false;
            }

            int matchMode = function == "XMATCH" || tokens.Count >= 3 && tokens[2].Length == 0 ? 0 : 1;
            FormulaArgumentValue matchError = default;
            if (tokens.Count >= 3 && tokens[2].Length > 0 && !TryGetLookupIntegerArgument(tokens[2], out matchMode, out matchError)) {
                return false;
            }

            if (!matchError.IsError && matchMode != -1 && matchMode != 0 && matchMode != 1
                && (function != "XMATCH" || matchMode != 2)) {
                return false;
            }

            int searchMode = 1;
            FormulaArgumentValue modeError = default;
            if (function == "XMATCH"
                && tokens.Count >= 4
                && tokens[3].Length > 0
                && (!TryGetLookupIntegerArgument(tokens[3], out searchMode, out modeError)
                    || (!modeError.IsError && searchMode != 1 && searchMode != -1 && searchMode != 2 && searchMode != -2))) {
                return false;
            }

            int budget = MaxResolvedFormulaRangeCells;
            if (!TryResolveLookupSource(tokens[1], ref budget, out FormulaLookupSource lookupValues)) return false;
            if (_formulaReferenceError.IsError) return true;
            if (lookupValues.Columns != 1 && lookupValues.Rows != 1) {
                result = FormulaArgumentValue.Error("#N/A");
                return true;
            }
            if (function == "MATCH") lookupValue = NormalizeLookupResult(lookupValue);
            int lookupMode = function == "MATCH" ? -matchMode : matchMode;
            if (function == "MATCH" && matchMode == 0 && !lookupValue.Number.HasValue) lookupMode = 2;
            if (lookupValue.IsError) { result = lookupValue; return true; }
            if (matchError.IsError) { result = matchError; return true; }
            if (modeError.IsError) { result = modeError; return true; }
            if (!TryFindLookupPosition(lookupValue, lookupValues, lookupMode, searchMode, out int position,
                out FormulaArgumentValue searchFailure, sortedApproximate: function == "MATCH" && matchMode != 0)) {
                if (searchFailure.IsUnresolvedFormula) return false;
                if (searchFailure.IsError) { result = searchFailure; return true; }
                result = FormulaArgumentValue.Error("#N/A");
                return true;
            }

            result = new FormulaArgumentValue(position, InvariantNumberText.Get(position));
            return true;
        }

        private bool TryFormatTextFunctionValue(FormulaArgumentValue value, string format, out string formatted) {
            formatted = string.Empty;
            const string englishLocale = "[$-409]";
            if (format.StartsWith(englishLocale, StringComparison.OrdinalIgnoreCase)) format = format.Substring(englishLocale.Length);
            if (format.IndexOf("[$-", StringComparison.Ordinal) >= 0) return false;
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
                    formatted = FormatFormulaDateText(value.Number.Value, date, dotNetFormat);
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
