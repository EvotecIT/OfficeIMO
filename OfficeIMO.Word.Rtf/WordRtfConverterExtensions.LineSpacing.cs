using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    // RTF encodes exact versus minimum point spacing in the sign. Automatic
    // spacing instead keeps its value in 240ths of a line, just as Word does.
    private static int? ToRtfLineSpacingValue(int? value, LineSpacingRuleValues? rule) =>
        rule == LineSpacingRuleValues.Exact && value > 0 ? -value : value;

    private static bool? ToRtfLineSpacingMultiple(int? value, LineSpacingRuleValues? rule) {
        if (!value.HasValue) return null;
        // An authored Word line value without a rule uses automatic units.
        return !rule.HasValue || rule.Value == LineSpacingRuleValues.Auto;
    }

    private static int? ToWordLineSpacingValue(int? value, bool? multiple) {
        if (multiple == true || value >= 0 || !value.HasValue) return value;
        if (value == int.MinValue) {
            throw new InvalidDataException("The RTF exact line-spacing magnitude exceeds the supported Word integer range.");
        }
        return -value;
    }

    private static LineSpacingRuleValues? ToWordLineSpacingRule(int? value, bool? multiple) {
        if (!value.HasValue) return null;
        if (multiple == true) return LineSpacingRuleValues.Auto;
        return value.Value < 0 ? LineSpacingRuleValues.Exact : LineSpacingRuleValues.AtLeast;
    }
}
