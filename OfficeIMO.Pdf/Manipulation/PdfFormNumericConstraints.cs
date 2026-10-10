using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

/// <summary>Recovers a small inert metadata grammar; never evaluates document scripts.</summary>
internal static class PdfFormNumericConstraints {
    private const string Number = @"[-+]?(?:[0-9]+(?:\.[0-9]+)?|\.[0-9]+)";
    private static readonly Regex Format = new(@"\A\s*AFNumber_(?<kind>Format|Keystroke)\s*\(\s*(?<precision>[0-9]{1,2})\s*,\s*(?<separator>[0-3])\s*,\s*0\s*,\s*0\s*,\s*(?:""""|'')\s*,\s*(?:true|false)\s*\)\s*;?\s*\z", RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(50));
    private static readonly Regex Range = new(@"\A\s*AFRange_Validate\s*\(\s*(?<hasMin>true|false)\s*,\s*(?<min>" + Number + @")\s*,\s*(?<hasMax>true|false)\s*,\s*(?<max>" + Number + @")\s*\)\s*;?\s*\z", RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(50));

    internal static IReadOnlyList<PdfFormFieldValueIssue> Assess(PdfFormField field, IReadOnlyList<string> values) {
        var actions = field.Actions.Concat(field.Widgets.SelectMany(widget => widget.Actions)).Where(action => action.IsJavaScript).ToArray();
        if (actions.Length == 0) return Array.Empty<PdfFormFieldValueIssue>();
        int? precision = null, separator = null; decimal? minimum = null, maximum = null;
        if (!field.IsTextField || actions.Length > 16) return Unsupported();
        foreach (var action in actions) {
            if (action.JavaScript is not { Length: <= 1024 } script) return Unsupported();
            Match format, range;
            try { format = Format.Match(script); range = Range.Match(script); }
            catch (RegexMatchTimeoutException) { return Unsupported(); }
            if (format.Success && (action.TriggerName == "K" && format.Groups["kind"].Value == "Keystroke" ||
                action.TriggerName == "F" && format.Groups["kind"].Value == "Format")) {
                int digits = int.Parse(format.Groups["precision"].Value, CultureInfo.InvariantCulture);
                int style = int.Parse(format.Groups["separator"].Value, CultureInfo.InvariantCulture);
                if (digits > 15 || precision.HasValue && precision != digits || separator.HasValue && separator != style) return Unsupported();
                precision = digits; separator = style;
            } else if (range.Success && action.TriggerName == "V") {
                if (!decimal.TryParse(range.Groups["min"].Value, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                    CultureInfo.InvariantCulture, out decimal min) || !decimal.TryParse(range.Groups["max"].Value,
                    NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal max)) return Unsupported();
                if (range.Groups["hasMin"].Value == "true") minimum = minimum.HasValue ? Math.Max(minimum.Value, min) : min;
                if (range.Groups["hasMax"].Value == "true") maximum = maximum.HasValue ? Math.Min(maximum.Value, max) : max;
                if (minimum.HasValue && maximum.HasValue && minimum > maximum) return Unsupported();
            } else return Unsupported();
        }
        var issues = new List<PdfFormFieldValueIssue>();
        foreach (string raw in values) {
            if (raw.Length == 0) continue;
            int style = separator ?? 1;
            char decimalSeparator = style >= 2 ? ',' : '.';
            char grouping = style == 0 ? ',' : style == 2 ? '.' : '\0';
            string text = raw.Trim();
            string[] pieces = text.Split(decimalSeparator);
            string integer = pieces[0].TrimStart('+', '-');
            bool valid = pieces.Length <= 2 && integer.Length > 0;
            if (grouping != '\0' && integer.Contains(grouping.ToString())) {
                string[] groups = integer.Split(grouping);
                valid &= groups[0].Length is >= 1 and <= 3 && groups.Skip(1).All(group => group.Length == 3) &&
                    groups.All(group => group.All(IsDigit));
            } else valid &= integer.All(IsDigit);
            valid &= pieces.Length == 1 || pieces[1].Length > 0 && pieces[1].All(IsDigit);
            string normalized = grouping == '\0' ? text : text.Replace(grouping.ToString(), string.Empty);
            normalized = normalized.Replace(decimalSeparator, '.');
            if (!valid || !decimal.TryParse(normalized, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture, out decimal numeric)) {
                issues.Add(new(PdfFormFieldValueIssueCode.NumericValue, true, "Enter a number using the field's declared decimal and grouping separators."));
                continue;
            }
            if (precision.HasValue && pieces.Length == 2 && pieces[1].Length > precision)
                issues.Add(new(PdfFormFieldValueIssueCode.NumericPrecision, true, $"Review policy requires at most {precision} decimal places. Correct the value before accepting it."));
            if (minimum.HasValue && numeric < minimum || maximum.HasValue && numeric > maximum)
                issues.Add(new(PdfFormFieldValueIssueCode.NumericRange, true, "The number is outside the field's declared range."));
        }
        return issues.AsReadOnly();
    }
    private static bool IsDigit(char value) => value is >= '0' and <= '9';
    private static PdfFormFieldValueIssue[] Unsupported() => new[] {
        new PdfFormFieldValueIssue(PdfFormFieldValueIssueCode.UnsupportedScriptConstraint, true,
            "This field has script constraints that cannot be safely assessed. Use a qualified form application.") };
}
