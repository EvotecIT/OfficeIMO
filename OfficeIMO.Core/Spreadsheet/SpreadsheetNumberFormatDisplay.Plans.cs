using System.Globalization;
using System.Text;

namespace OfficeIMO.Spreadsheet {
    internal static partial class SpreadsheetNumberFormatDisplay {
        /// <summary>
        /// Splits a format code into its sections once. Returns null when a section carries a condition such as
        /// <c>[&gt;100]</c>, because conditional selection depends on each value and stays with the per-call path.
        /// </summary>
        internal static string[]? SplitUnconditionalSections(string formatCode) {
            string[] sections = SplitNumberFormatSections(formatCode);
            foreach (string section in sections) {
                if (TryGetSectionCondition(section, out _, out _)) return null;
            }
            return sections.Length == 0 ? new[] { formatCode } : sections;
        }

        /// <summary>
        /// Finds a <c>General</c> placeholder outside quoted text, as in <c>General;[Red]General</c> or
        /// <c>"Total: "General</c>, where the section shows the value in General form.
        /// </summary>
        internal static bool TryFindGeneralToken(string section, out int start) {
            start = -1;
            bool inQuote = false;
            for (int index = 0; index + 7 <= section.Length; index++) {
                char ch = section[index];
                if (ch == '"') { inQuote = !inQuote; continue; }
                if (inQuote) continue;
                if (ch == '\\' || ch == '_' || ch == '*') { index++; continue; }
                if (ch == '[') {
                    int close = section.IndexOf(']', index + 1);
                    if (close > index) { index = close; continue; }
                }
                if (string.Compare(section, index, "General", 0, 7, StringComparison.OrdinalIgnoreCase) == 0) {
                    start = index;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Shows the value in General form between the section's literal affixes. A negative section shows the
        /// magnitude, because the section itself marks the sign, as numeric placeholders do.
        /// </summary>
        internal static string FormatGeneralSection(double value, string section, int generalStart, int selectedSection, IFormatProvider provider) {
            string prefix = CleanLiteralAffix(StripNumberFormatDecorations(section.Substring(0, generalStart)));
            string suffix = CleanLiteralAffix(StripNumberFormatDecorations(section.Substring(generalStart + 7)));
            return prefix + FormatGeneralValue(selectedSection == 1 ? Math.Abs(value) : value, provider) + suffix;
        }

        /// <summary>Selects a section the same way the per-call formatter does for unconditional formats.</summary>
        internal static int SelectSectionIndex(double value, int sectionCount) {
            int preferred = value < 0 ? 1 : value == 0 ? 2 : 0;
            return preferred < sectionCount ? preferred : 0;
        }

        /// <summary>
        /// A numeric format section analysed once. <see cref="Format"/> produces the same text as
        /// <see cref="FormatNumericValue"/> for that section without re-parsing it.
        /// </summary>
        internal sealed class NumericSectionPlan {
            private readonly bool _empty;
            private readonly string _section;
            private readonly int _generalStart;
            private readonly string _normalized;
            private readonly bool _textPlaceholder;
            private readonly string? _literal;
            private readonly bool _literalOnly;
            private readonly bool _onlyOptional;
            private readonly bool _fraction;
            private readonly string? _scientificCode;
            private readonly int _percent;
            private readonly int _scaling;
            private readonly bool _thousands;
            private readonly DecimalPlaceInfo _decimals;
            private readonly string _numericFormat;
            private readonly int _requiredIntegerDigits;
            private readonly string _prefix;
            private readonly string _suffix;
            private readonly bool _currencyPrefix;

            private NumericSectionPlan(string section) {
                _empty = section.Length == 0;
                _section = section;
                _generalStart = TryFindGeneralToken(section, out int generalStart) ? generalStart : -1;
                _normalized = StripNumberFormatDecorations(section);
                string lower = _normalized.ToLowerInvariant();
                _textPlaceholder = SpreadsheetNumberFormatSyntax.Parse(section).Tokens.Any(token => token.Kind == SpreadsheetNumberFormatTokenKind.TextPlaceholder);
                _literalOnly = !ContainsNumericPlaceholder(_normalized);
                if (_literalOnly) {
                    string literal = CleanLiteralAffix(_normalized);
                    _literal = string.IsNullOrEmpty(literal) ? null : literal;
                }
                _onlyOptional = HasOnlyOptionalDigitPlaceholders(_normalized);
                _fraction = _normalized.IndexOf('/') >= 0;
                _scientificCode = CompileScientific(_normalized);
                _percent = CountPercentPlaceholders(section);
                _thousands = lower.Contains("#,##") || lower.Contains(",##");
                _decimals = GetDecimalPlaceInfo(lower);
                _scaling = CountScalingCommas(_normalized);
                _numericFormat = (_thousands ? "N" : "F") + _decimals.Maximum.ToString(CultureInfo.InvariantCulture);
                _requiredIntegerDigits = GetRequiredIntegerDigits(_normalized);
                int first = FindFirstNumericPlaceholder(_normalized);
                int last = FindLastNumericPlaceholder(_normalized);
                if (first >= 0 && last >= first) {
                    _prefix = CleanLiteralAffix(_normalized.Substring(0, first));
                    _suffix = CleanLiteralAffix(_normalized.Substring(last + 1));
                } else {
                    _prefix = CleanLiteralAffix(_normalized);
                    _suffix = string.Empty;
                }
                _currencyPrefix = _prefix.Length > 0 && _prefix[0] is '$' or '€' or '£';
            }

            internal static NumericSectionPlan Compile(string section) => new(section);

            /// <summary>Formats a value for this section; null means the caller's General fallback applies.</summary>
            internal string? Format(double value, int selectedSection, IFormatProvider provider) {
                if (double.IsNaN(value) || double.IsInfinity(value)) return null;
                if (_empty) return string.Empty;
                if (_generalStart >= 0) return FormatGeneralSection(value, _section, _generalStart, selectedSection, provider);
                if (_textPlaceholder) return value.ToString(provider);
                if (_literalOnly) return _literal;
                if (IsZeroValue(value) && _onlyOptional) return string.Empty;
                if (_fraction && TryFormatFraction(value, _normalized, selectedSection, out string fractionText)) {
                    return ApplyFractionAffixes(_normalized, fractionText);
                }
                if (_scientificCode != null) {
                    double scientificValue = selectedSection == 1 ? Math.Abs(value) : value;
                    return Affix(scientificValue.ToString(_scientificCode, provider));
                }

                double displayValue = value;
                for (int i = 0; i < _percent; i++) displayValue *= 100D;
                for (int i = 0; i < _scaling; i++) displayValue /= 1000D;
                if (double.IsNaN(displayValue) || double.IsInfinity(displayValue)) return null;
                if (selectedSection == 1) displayValue = Math.Abs(displayValue);
                double rounded = _decimals.Maximum <= 15
                    ? Math.Round(displayValue, _decimals.Maximum, MidpointRounding.AwayFromZero)
                    : displayValue;
                var numberFormat = NumberFormatInfo.GetInstance(provider);
                if (_decimals.Optional == 0 && _requiredIntegerDigits <= 1) {
                    return AffixFormatted(rounded, numberFormat);
                }
                string text = rounded.ToString(_numericFormat, numberFormat);
                if (_decimals.Optional > 0) text = TrimOptionalDecimalPlaces(text, _decimals.Required, numberFormat);
                text = PadRequiredIntegerDigits(text, _requiredIntegerDigits, _thousands, numberFormat);
                return Affix(text);
            }

            private string Affix(string numericText) {
                if (_currencyPrefix && numericText.Length > 0 && numericText[0] == '-') {
                    return "-" + _prefix + numericText.Substring(1) + _suffix;
                }
                return _prefix + numericText + _suffix;
            }

#if NET8_0_OR_GREATER
            // Formats into a stack buffer and allocates only the final string.
            private string AffixFormatted(double rounded, NumberFormatInfo numberFormat) {
                Span<char> buffer = stackalloc char[128];
                if (!rounded.TryFormat(buffer, out int written, _numericFormat, numberFormat)) {
                    return Affix(rounded.ToString(_numericFormat, numberFormat));
                }
                ReadOnlySpan<char> number = buffer.Slice(0, written);
                if (_currencyPrefix && number.Length > 0 && number[0] == '-') {
                    return string.Concat("-", _prefix, string.Concat(number.Slice(1), _suffix.AsSpan()));
                }
                return string.Concat(_prefix.AsSpan(), number, _suffix.AsSpan());
            }
#else
            private string AffixFormatted(double rounded, NumberFormatInfo numberFormat) =>
                Affix(rounded.ToString(_numericFormat, numberFormat));
#endif

            /// <summary>Precomputes the .NET pattern the scientific formatter would build for every value.</summary>
            private static string? CompileScientific(string formatCode) {
                int first = FindFirstNumericPlaceholder(formatCode);
                if (first < 0) return null;
                for (int index = first; index < formatCode.Length; index++) {
                    char token = formatCode[index];
                    if (token == LiteralPunctuationMarker && index + 1 < formatCode.Length) { index++; continue; }
                    if (token is not 'E' and not 'e' || index + 2 >= formatCode.Length) continue;
                    char sign = formatCode[index + 1];
                    if (sign is not '+' and not '-') continue;
                    int digits = 0;
                    while (index + 2 + digits < formatCode.Length && formatCode[index + 2 + digits] == '0') digits++;
                    if (digits == 0) continue;
                    DecimalPlaceInfo precision = GetDecimalPlaceInfo(formatCode.Substring(first, index - first));
                    string mantissa = "0" + (precision.Maximum == 0 ? string.Empty
                        : "." + new string('0', precision.Required) + new string('#', precision.Optional));
                    return mantissa + token + sign + new string('0', digits);
                }
                return null;
            }
        }

        /// <summary>
        /// A date/time format section analysed once. <see cref="Format"/> produces the same text as
        /// <see cref="FormatDateTimeValue"/> for that section. Elapsed-time sections are not compiled.
        /// </summary>
        internal sealed class TemporalSectionPlan {
            private enum OpKind : byte { Pattern, MonthInitial, Meridiem }

            private readonly struct Op {
                internal Op(OpKind kind, string text, string? second = null) { Kind = kind; Text = text; Second = second; }
                internal OpKind Kind { get; }
                internal string Text { get; }
                internal string? Second { get; }
            }

            private readonly IReadOnlyList<SpreadsheetNumberFormatToken> _tokens;
            private readonly Op[] _ops;
            private readonly IFormatProvider _provider;
            private readonly bool _requiresCalendar;
            private readonly DateTime _minimum;
            private readonly DateTime _maximum;

            private TemporalSectionPlan(IReadOnlyList<SpreadsheetNumberFormatToken> tokens, Op[] ops, IFormatProvider provider, bool requiresCalendar) {
                _tokens = tokens;
                _ops = ops;
                _provider = provider;
                _requiresCalendar = requiresCalendar;
                Calendar calendar = DateTimeFormatInfo.GetInstance(provider).Calendar;
                _minimum = calendar.MinSupportedDateTime;
                _maximum = calendar.MaxSupportedDateTime;
            }

            /// <summary>Returns null when the section uses syntax the plan does not cover; callers then use the per-call path.</summary>
            internal static TemporalSectionPlan? Compile(string section, IFormatProvider provider) {
                var syntax = SpreadsheetNumberFormatSyntax.Parse(section);
                if (!syntax.IsValid || syntax.Tokens.Any(token => ElapsedUnit(token) != '\0')) return null;
                provider = ResolveTemporalProvider(syntax, provider);
                bool twelveHour = syntax.Tokens.Select((_, index) => index).Any(index => TryReadMeridiem(syntax, index, out _, out _));
                var ops = new List<Op>();
                var pattern = new StringBuilder();
                int patternParts = 0;
                string? singlePart = null;
                void FlushPattern() {
                    if (patternParts == 0) return;
                    ops.Add(new Op(OpKind.Pattern, patternParts == 1 && singlePart != null ? singlePart : pattern.ToString()));
                    pattern.Clear();
                    patternParts = 0;
                    singlePart = null;
                }
                void AddIsolated(string specifier) {
                    // .NET switches month names to their genitive form when a day shares the pattern,
                    // so month names are rendered on their own, as the per-call formatter does.
                    FlushPattern();
                    ops.Add(new Op(OpKind.Pattern, specifier));
                }
                void AddPattern(string specifier) {
                    // A one-letter custom specifier needs '%' alone; inside a longer pattern it does not.
                    singlePart = patternParts == 0 ? specifier : null;
                    pattern.Append(specifier[0] == '%' ? specifier.Substring(1) : specifier);
                    patternParts++;
                }
                void AddLiteral(string literal) {
                    foreach (char ch in literal) pattern.Append('\\').Append(ch);
                    singlePart = null;
                    patternParts += 2;
                }

                var literalBuilder = new StringBuilder();
                for (int index = 0; index < syntax.Tokens.Count; index++) {
                    var token = syntax.Tokens[index];
                    if (TryReadMeridiem(syntax, index, out string meridiem, out int lastIndex)) {
                        FlushPattern();
                        string[] halves = meridiem.Split('/');
                        ops.Add(new Op(OpKind.Meridiem, halves[0], halves[1]));
                        index = lastIndex;
                    } else if (token.Kind == SpreadsheetNumberFormatTokenKind.DateTimeSymbol) {
                        char unit = char.ToLowerInvariant(token.Text[0]);
                        int length = token.Text.Length;
                        switch (unit) {
                            case 'y': AddPattern(length <= 2 ? "yy" : "yyyy"); break;
                            case 'd': AddPattern(length >= 4 ? "dddd" : length == 3 ? "ddd" : length == 2 ? "dd" : "%d"); break;
                            case 'm':
                                if (length == 5) {
                                    FlushPattern();
                                    ops.Add(new Op(OpKind.MonthInitial, "MMMM"));
                                    break;
                                }
                                if (length >= 3) {
                                    AddIsolated(length >= 4 ? "MMMM" : "MMM");
                                    break;
                                }
                                AddPattern(IsTemporalMinute(syntax.Tokens, index) ? length == 1 ? "%m" : "mm"
                                    : length == 1 ? "%M" : "MM");
                                break;
                            case 'h': AddPattern(twelveHour ? length == 1 ? "%h" : "hh" : length == 1 ? "%H" : "HH"); break;
                            case 's': AddPattern(length == 1 ? "%s" : "ss"); break;
                            default: return null;
                        }
                    } else if (IsFractionalSecond(syntax.Tokens, index)) {
                        if (token.Text.Length > 7) return null;
                        AddPattern(token.Text.Length == 1 ? "%f" : new string('f', token.Text.Length));
                    } else {
                        literalBuilder.Clear();
                        if (!AppendTemporalLiteral(literalBuilder, token)) return null;
                        if (literalBuilder.Length > 0) AddLiteral(literalBuilder.ToString());
                    }
                }
                FlushPattern();
                return new TemporalSectionPlan(syntax.Tokens, ops.ToArray(), provider, RequiresCalendarDate(syntax));
            }

            internal string? Format(DateTime date) {
                if (!TryRoundTemporalTicks(date.Ticks, _tokens, out long roundedTicks)) return null;
                try {
                    date = new DateTime(roundedTicks, date.Kind);
                } catch (ArgumentOutOfRangeException) {
                    return null;
                }
                if (_requiresCalendar && (date < _minimum || date > _maximum)) return null;
                if (_ops.Length == 1 && _ops[0].Kind == OpKind.Pattern) return date.ToString(_ops[0].Text, _provider);
                var builder = new StringBuilder(32);
                foreach (Op op in _ops) {
                    switch (op.Kind) {
                        case OpKind.Pattern: builder.Append(date.ToString(op.Text, _provider)); break;
                        case OpKind.MonthInitial: builder.Append(date.ToString(op.Text, _provider)[0]); break;
                        default: builder.Append(date.Hour < 12 ? op.Text : op.Second); break;
                    }
                }
                return builder.ToString();
            }
        }
    }
}
