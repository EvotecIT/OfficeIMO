using System.Globalization;
using OfficeIMO.Spreadsheet;
using static OfficeIMO.Spreadsheet.SpreadsheetNumberFormatDisplay;

namespace OfficeIMO.Excel {
    /// <summary>
    /// A cell number format parsed once for repeated display formatting. Its output matches
    /// <see cref="ExcelNumberFormatDisplay.FormatNumericText"/>; conditional and elapsed-time formats, which depend on
    /// each value's section choice, keep that per-call path.
    /// </summary>
    internal sealed class ExcelNumberFormatPlan {
        private readonly ExcelDateSystem _dateSystem;
        private readonly IFormatProvider _provider;
        private readonly Section[]? _sections;

        private sealed class Section {
            internal NumericSectionPlan? Numeric;
            internal TemporalSectionPlan? Temporal;
            internal bool FractionalSeconds;
        }

        private ExcelNumberFormatPlan(uint numberFormatId, string formatCode, ExcelDateSystem dateSystem, IFormatProvider provider) {
            NumberFormatId = numberFormatId;
            FormatCode = formatCode;
            _dateSystem = dateSystem;
            _provider = provider;
            _sections = Compile(numberFormatId, formatCode, provider);
        }

        internal uint NumberFormatId { get; }

        internal string FormatCode { get; }

        /// <summary>Returns null for formats that display like General, so callers keep their typed-value route.</summary>
        internal static ExcelNumberFormatPlan? Create(uint numberFormatId, string? formatCode, ExcelDateSystem dateSystem, IFormatProvider provider) {
            if (numberFormatId == 0U || numberFormatId == 49U || string.IsNullOrWhiteSpace(formatCode)
                || string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase) || formatCode == "@") {
                return null;
            }
            return new ExcelNumberFormatPlan(numberFormatId, formatCode!, dateSystem, provider);
        }

        private static Section[]? Compile(uint numberFormatId, string formatCode, IFormatProvider provider) {
            string[]? sections = SplitUnconditionalSections(formatCode);
            if (sections == null) return null;
            var compiled = new Section[sections.Length];
            for (int index = 0; index < sections.Length; index++) {
                string section = sections[index];
                if (ExcelNumberFormatDisplay.IsDateNumberFormat(numberFormatId, section)) {
                    if (HasElapsedTimeToken(section)) return null;
                    TemporalSectionPlan? temporal = TemporalSectionPlan.Compile(section, provider);
                    if (temporal == null) return null;
                    compiled[index] = new Section { Temporal = temporal, FractionalSeconds = HasFractionalSecondToken(section) };
                } else {
                    compiled[index] = new Section { Numeric = NumericSectionPlan.Compile(section) };
                }
            }
            return compiled;
        }

        internal string Format(double value) {
            if (_sections == null) return FormatPerCall(value);
            int selected = SelectSectionIndex(value, _sections.Length);
            Section section = _sections[selected];
            if (section.Numeric != null) {
                return section.Numeric.Format(value, selected, _provider) ?? General(value);
            }

            DateTime date;
            try {
                date = section.FractionalSeconds
                    ? ExcelDateSystemConverter.FromSerialForDisplay(value, _dateSystem)
                    : ExcelDateSystemConverter.FromSerial(value, _dateSystem);
            } catch (ArgumentException) {
                return value.ToString(CultureInfo.InvariantCulture);
            }
            string? text;
            if (NumberFormatId == 14U) {
                Calendar calendar = DateTimeFormatInfo.GetInstance(_provider).Calendar;
                text = date < calendar.MinSupportedDateTime || date > calendar.MaxSupportedDateTime ? null : date.ToString("d", _provider);
            } else {
                text = section.Temporal!.Format(date);
            }
            return text ?? value.ToString(_provider);
        }

        private string General(double value) => FormatGeneralValue(value, _provider);

        private string FormatPerCall(double value) =>
            ExcelNumberFormatDisplay.FormatNumericText(value, NumberFormatId, FormatCode, General(value), _dateSystem, _provider);
    }
}
