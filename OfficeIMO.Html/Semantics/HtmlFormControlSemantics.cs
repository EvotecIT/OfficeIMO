using AngleSharp.Dom;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

/// <summary>Canonical HTML form-control normalization shared by semantic projection and fidelity scoring.</summary>
internal static class HtmlFormControlSemantics {
    internal static string GetEffectiveType(string elementName, string? type, bool multiple = false) {
        string name = NormalizeIdentifier(elementName);
        string normalized = NormalizeKeyword(type);
        if (name == "input") return IsValidInputType(normalized) ? normalized : "text";
        if (name == "button") return IsValidButtonType(normalized) ? normalized : "submit";
        if (name == "select") return multiple ? "select-multiple" : "select-one";
        return name;
    }

    internal static bool IsValidType(string elementName, string? type) {
        string name = NormalizeIdentifier(elementName);
        string normalized = NormalizeKeyword(type);
        return name == "input" ? IsValidInputType(normalized)
            : name == "button" && IsValidButtonType(normalized);
    }

    internal static string GetEffectiveFormMethod(string? method) {
        string normalized = NormalizeKeyword(method);
        return normalized == "post" || normalized == "dialog" ? normalized : "get";
    }

    internal static string GetEffectiveFormEncoding(string? encodingType) {
        string normalized = NormalizeKeyword(encodingType);
        return normalized == "multipart/form-data" || normalized == "text/plain"
            ? normalized
            : "application/x-www-form-urlencoded";
    }

    internal static bool IsSubmitter(IElement element) {
        string name = NormalizeIdentifier(element.LocalName);
        string type = GetEffectiveType(name, element.GetAttribute("type"));
        return name == "button" ? type != "button" && type != "reset"
            : name == "input" && (type == "submit" || type == "image");
    }

    internal static bool IsCheckedStateApplicable(string elementName, string effectiveType) =>
        NormalizeIdentifier(elementName) == "input" && (effectiveType == "checkbox" || effectiveType == "radio");

    internal static bool IsMultipleStateApplicable(string elementName, string effectiveType) {
        string name = NormalizeIdentifier(elementName);
        return name == "select" || name == "input" && (effectiveType == "email" || effectiveType == "file");
    }

    internal static bool IsRequiredStateApplicable(string elementName, string effectiveType) {
        string name = NormalizeIdentifier(elementName);
        if (name == "select" || name == "textarea") return true;
        if (name != "input") return false;
        switch (effectiveType) {
            case "checkbox":
            case "date":
            case "datetime-local":
            case "email":
            case "file":
            case "month":
            case "number":
            case "password":
            case "radio":
            case "search":
            case "tel":
            case "text":
            case "time":
            case "url":
            case "week":
                return true;
            default:
                return false;
        }
    }

    internal static bool IsReadOnlyStateApplicable(string elementName, string effectiveType) {
        string name = NormalizeIdentifier(elementName);
        if (name == "textarea") return true;
        if (name != "input") return false;
        switch (effectiveType) {
            case "date":
            case "datetime-local":
            case "email":
            case "month":
            case "number":
            case "password":
            case "search":
            case "tel":
            case "text":
            case "time":
            case "url":
            case "week":
                return true;
            default:
                return false;
        }
    }

    internal static bool IsPatternApplicable(string elementName, string effectiveType) =>
        NormalizeIdentifier(elementName) == "input" && IsTextInputType(effectiveType);

    internal static bool IsLengthApplicable(string elementName, string effectiveType) =>
        NormalizeIdentifier(elementName) == "textarea"
        || NormalizeIdentifier(elementName) == "input" && IsTextInputType(effectiveType);

    internal static bool TryParseLengthConstraint(string? value, out int length) =>
        HtmlIntegerSemantics.TryParseNonNegativeInteger(value, out length);

    internal static bool IsRangeApplicable(string elementName, string effectiveType) {
        if (NormalizeIdentifier(elementName) != "input") return false;
        switch (effectiveType) {
            case "date":
            case "datetime-local":
            case "month":
            case "number":
            case "range":
            case "time":
            case "week":
                return true;
            default:
                return false;
        }
    }

    internal static bool IsPlaceholderApplicable(string elementName, string effectiveType) =>
        NormalizeIdentifier(elementName) == "textarea"
        || NormalizeIdentifier(elementName) == "input" && IsTextInputType(effectiveType);

    internal static bool IsStateAttributeApplicable(string elementName, string effectiveType, string attributeName) {
        string name = NormalizeIdentifier(elementName);
        switch (NormalizeIdentifier(attributeName)) {
            case "type": return name == "input" || name == "button";
            case "name": return name == "input" || name == "select" || name == "textarea" || name == "button";
            case "checked": return IsCheckedStateApplicable(elementName, effectiveType);
            case "selected": return name == "option";
            case "disabled": return name == "input" || name == "select" || name == "textarea"
                || name == "button" || name == "option";
            case "multiple": return IsMultipleStateApplicable(elementName, effectiveType);
            case "required": return IsRequiredStateApplicable(elementName, effectiveType);
            case "readonly": return IsReadOnlyStateApplicable(elementName, effectiveType);
            case "pattern": return IsPatternApplicable(elementName, effectiveType);
            case "minlength":
            case "maxlength": return IsLengthApplicable(elementName, effectiveType);
            case "min":
            case "max":
            case "step": return IsRangeApplicable(elementName, effectiveType);
            case "placeholder": return IsPlaceholderApplicable(elementName, effectiveType);
            case "value": return name == "button" || name == "option"
                || name == "input" && effectiveType != "file";
            case "autocomplete": return name == "input" || name == "select" || name == "textarea";
            case "data-fieldset-disabled": return name == "input" || name == "select"
                || name == "textarea" || name == "button";
            default: return true;
        }
    }

    internal static IReadOnlyList<string> GetValues(IElement element) {
        string name = NormalizeIdentifier(element.LocalName);
        if (name == "select") return GetSelectValues(element);
        if (name == "textarea") return new[] { element.TextContent ?? string.Empty };
        if (name == "input") {
            string effectiveType = GetEffectiveType(name, element.GetAttribute("type"));
            if (effectiveType == "file") return Array.Empty<string>();
            if (effectiveType == "range") {
                return new[] { GetRangeValue(
                    element.GetAttribute("value"),
                    element.GetAttribute("min"),
                    element.GetAttribute("max"),
                    element.GetAttribute("step")) };
            }
            string value = element.GetAttribute("value")
                ?? GetDefaultValue(name, element.GetAttribute("type"), element.TextContent ?? string.Empty);
            return new[] { SanitizeInputValue(effectiveType, value, element.HasAttribute("multiple")) };
        }
        if (name == "button") return new[] { element.GetAttribute("value") ?? string.Empty };
        return Array.Empty<string>();
    }

    internal static string GetDefaultValue(string elementName, string? type, string textContent) {
        string name = NormalizeIdentifier(elementName);
        if (name == "option") return NormalizeOptionText(textContent);
        if (name == "input") {
            string effectiveType = GetEffectiveType(name, type);
            if (effectiveType == "checkbox" || effectiveType == "radio") return "on";
            if (effectiveType == "color") return "#000000";
        }
        return string.Empty;
    }

    internal static string GetOptionLabel(IElement option) =>
        option.GetAttribute("label") ?? NormalizeOptionText(option.TextContent);

    internal static bool IsEffectivelyDisabled(IElement element) {
        if (element.HasAttribute("disabled")) return true;
        for (IElement? ancestor = element.ParentElement; ancestor != null; ancestor = ancestor.ParentElement) {
            if (!string.Equals(ancestor.LocalName, "fieldset", StringComparison.OrdinalIgnoreCase)
                || !ancestor.HasAttribute("disabled")) continue;

            IElement? firstLegend = ancestor.Children.FirstOrDefault(child =>
                string.Equals(child.LocalName, "legend", StringComparison.OrdinalIgnoreCase));
            if (firstLegend == null || !IsDescendantOf(element, firstLegend)) return true;
        }
        return false;
    }

    internal static string ResolveFormOwnerId(IElement element) {
        IElement? owner = ResolveFormOwner(element);
        return owner?.GetAttribute("id") ?? string.Empty;
    }

    internal static IElement? ResolveFormOwner(IElement element) {
        if (element.HasAttribute("form")) {
            string explicitOwner = element.GetAttribute("form") ?? string.Empty;
            if (explicitOwner.Length == 0) return null;
            IElement? candidate = element.Owner?.GetElementById(explicitOwner);
            return candidate != null && string.Equals(candidate.LocalName, "form", StringComparison.OrdinalIgnoreCase)
                ? candidate
                : null;
        }

        for (IElement? ancestor = element.ParentElement; ancestor != null; ancestor = ancestor.ParentElement) {
            if (string.Equals(ancestor.LocalName, "form", StringComparison.OrdinalIgnoreCase)) {
                return ancestor;
            }
        }
        return null;
    }

    internal static bool IsEffectivelyChecked(IElement element) {
        string effectiveType = GetEffectiveType(element.LocalName, element.GetAttribute("type"));
        if (!IsCheckedStateApplicable(element.LocalName, effectiveType) || !element.HasAttribute("checked")) {
            return false;
        }
        if (effectiveType == "checkbox") return true;

        string name = element.GetAttribute("name") ?? string.Empty;
        if (name.Length == 0 || element.Owner == null) return true;
        IElement? owner = ResolveFormOwner(element);
        IElement? effective = element.Owner.QuerySelectorAll("input[checked]").LastOrDefault(candidate =>
            string.Equals(GetEffectiveType(candidate.LocalName, candidate.GetAttribute("type")), "radio", StringComparison.Ordinal)
            && string.Equals(candidate.GetAttribute("name") ?? string.Empty, name, StringComparison.Ordinal)
            && ReferenceEquals(ResolveFormOwner(candidate), owner));
        return ReferenceEquals(element, effective);
    }

    private static IReadOnlyList<string> GetSelectValues(IElement select) {
        return GetEffectiveSelectedOptions(select).Select(GetOptionValue).ToArray();
    }

    internal static IReadOnlyList<IElement> GetEffectiveSelectedOptions(IElement select) {
        IElement[] options = select.QuerySelectorAll("option").ToArray();
        IElement[] selected = options.Where(option => option.HasAttribute("selected")).ToArray();
        if (select.HasAttribute("multiple")) return selected;
        IElement? effective = selected.LastOrDefault();
        if (effective == null && GetSelectDisplaySize(select) == 1) {
            effective = options.FirstOrDefault(option => !IsOptionEffectivelyDisabled(option));
        }
        return effective == null ? Array.Empty<IElement>() : new[] { effective };
    }

    internal static int GetSelectDisplaySize(IElement select) {
        if (select.HasAttribute("size")
            && HtmlIntegerSemantics.TryParseNonNegativeInteger(select.GetAttribute("size"), out int size)
            && size > 0) {
            return size;
        }
        return select.HasAttribute("multiple") ? 4 : 1;
    }

    internal static bool IsOptionEffectivelyDisabled(IElement option) {
        if (!string.Equals(option.LocalName, "option", StringComparison.OrdinalIgnoreCase)) return false;
        if (option.HasAttribute("disabled")) return true;
        for (IElement? ancestor = option.ParentElement; ancestor != null; ancestor = ancestor.ParentElement) {
            if (string.Equals(ancestor.LocalName, "optgroup", StringComparison.OrdinalIgnoreCase)
                && ancestor.HasAttribute("disabled")) return true;
            if (string.Equals(ancestor.LocalName, "select", StringComparison.OrdinalIgnoreCase)) break;
        }
        return false;
    }

    internal static string GetOptionValue(IElement option) =>
        option.GetAttribute("value") ?? GetDefaultValue("option", null, option.TextContent ?? string.Empty);

    internal static string GetRangeValue(string? value, string? minimum, string? maximum, string? step = null) =>
        ResolveRange(value, minimum, maximum, step).ValueText;

    internal static double GetRangeFraction(IElement element) => ResolveRange(
        element.GetAttribute("value"),
        element.GetAttribute("min"),
        element.GetAttribute("max"),
        element.GetAttribute("step")).Fraction;

    private static HtmlRangeState ResolveRange(string? value, string? minimum, string? maximum, string? step) {
        bool hasMinimum = TryParseHtmlNumber(minimum, out double parsedMinimum);
        double min = hasMinimum ? parsedMinimum : 0D;
        double max = TryParseHtmlNumber(maximum, out double parsedMaximum) ? parsedMaximum : 100D;
        if (max < min) return new HtmlRangeState(FormatHtmlNumber(min), 0D);

        bool hasAuthoredValue = TryParseHtmlNumber(value, out double current);
        bool preservesAuthoredValue = hasAuthoredValue;
        double authoredValue = current;
        if (!preservesAuthoredValue) current = Midpoint(min, max);
        if (current < min) {
            current = min;
            preservesAuthoredValue = false;
        } else if (current > max) {
            current = max;
            preservesAuthoredValue = false;
        }

        if (!string.Equals(NormalizeKeyword(step), "any", StringComparison.Ordinal)) {
            double allowedStep = TryParseHtmlNumber(step, out double parsedStep) && parsedStep > 0D ? parsedStep : 1D;
            double stepBase = hasMinimum ? min : hasAuthoredValue ? authoredValue : 0D;
            double stepped = RoundToAllowedStep(current, min, max, stepBase, allowedStep);
            if (stepped != current) {
                current = stepped;
                preservesAuthoredValue = false;
            }
        }

        double fraction = ResolveRangeFraction(current, min, max);
        return new HtmlRangeState(
            preservesAuthoredValue ? value! : FormatHtmlNumber(current),
            fraction);
    }

    private static double RoundToAllowedStep(double value, double minimum, double maximum, double stepBase, double step) {
        double quotient = (value - stepBase) / step;
        if (double.IsNaN(quotient) || double.IsInfinity(quotient)) return value;
        double nearestIndex = Math.Floor(quotient + 0.5D);
        double candidate = stepBase + nearestIndex * step;
        if (double.IsNaN(candidate) || double.IsInfinity(candidate)) return value;
        if (candidate < minimum) candidate += step;
        if (candidate > maximum) candidate -= step;
        if (candidate < minimum || candidate > maximum) return value;
        double tolerance = 1e-12D * Math.Max(1D, Math.Max(Math.Abs(value), Math.Abs(candidate)));
        return Math.Abs(candidate - value) <= tolerance ? value : candidate;
    }

    private static bool TryParseHtmlNumber(string? text, out double value) {
        string candidate = text ?? string.Empty;
        if (!Regex.IsMatch(candidate, @"^-?(?:[0-9]+(?:\.[0-9]+)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$", RegexOptions.CultureInvariant)
            || !double.TryParse(candidate, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
            || double.IsNaN(value)
            || double.IsInfinity(value)) {
            value = 0D;
            return false;
        }
        return true;
    }

    private static string SanitizeInputValue(string effectiveType, string value, bool multiple) {
        switch (effectiveType) {
            case "text":
            case "search":
            case "tel":
            case "password":
                return StripLineBreaks(value);
            case "url":
                return TrimAsciiWhitespace(StripLineBreaks(value));
            case "email":
                string email = StripLineBreaks(value);
                return multiple
                    ? string.Join(",", email.Split(',').Select(TrimAsciiWhitespace))
                    : TrimAsciiWhitespace(email);
            case "date":
                return IsValidDate(value) ? value : string.Empty;
            case "month":
                return IsValidMonth(value) ? value : string.Empty;
            case "week":
                return IsValidWeek(value) ? value : string.Empty;
            case "time":
                return IsValidTime(value) ? value : string.Empty;
            case "datetime-local":
                return NormalizeLocalDateAndTime(value);
            case "number":
                return TryParseHtmlNumber(value, out _) ? value : string.Empty;
            case "color":
                return Regex.IsMatch(value, "^#[0-9A-Fa-f]{6}$", RegexOptions.CultureInvariant)
                    ? value.ToLowerInvariant()
                    : "#000000";
            default:
                return value;
        }
    }

    private static bool IsValidDate(string value) {
        Match match = Regex.Match(value, "^([0-9]{4,})-([0-9]{2})-([0-9]{2})$", RegexOptions.CultureInvariant);
        if (!match.Success || !TryReadPositiveYear(match.Groups[1].Value, out int yearModulo400)
            || !int.TryParse(match.Groups[2].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int month)
            || !int.TryParse(match.Groups[3].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int day)) {
            return false;
        }
        int maximumDay = DaysInMonth(yearModulo400, month);
        return maximumDay > 0 && day >= 1 && day <= maximumDay;
    }

    private static bool IsValidMonth(string value) {
        Match match = Regex.Match(value, "^([0-9]{4,})-([0-9]{2})$", RegexOptions.CultureInvariant);
        return match.Success
            && TryReadPositiveYear(match.Groups[1].Value, out _)
            && int.TryParse(match.Groups[2].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int month)
            && month >= 1
            && month <= 12;
    }

    private static bool IsValidWeek(string value) {
        Match match = Regex.Match(value, "^([0-9]{4,})-W([0-9]{2})$", RegexOptions.CultureInvariant);
        if (!match.Success || !TryReadPositiveYear(match.Groups[1].Value, out int yearModulo400)
            || !int.TryParse(match.Groups[2].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int week)) {
            return false;
        }
        return week >= 1 && week <= WeeksInYear(yearModulo400);
    }

    private static bool IsValidTime(string value) {
        Match match = Regex.Match(
            value,
            "^([0-9]{2}):([0-9]{2})(?::([0-9]{2})(?:\\.([0-9]+))?)?$",
            RegexOptions.CultureInvariant);
        return match.Success
            && int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture) <= 23
            && int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) <= 59
            && (!match.Groups[3].Success || int.Parse(match.Groups[3].Value, CultureInfo.InvariantCulture) <= 59);
    }

    private static string NormalizeLocalDateAndTime(string value) {
        Match match = Regex.Match(value, "^(.+?)[T ](.+)$", RegexOptions.CultureInvariant);
        if (!match.Success || !IsValidDate(match.Groups[1].Value) || !IsValidTime(match.Groups[2].Value)) {
            return string.Empty;
        }

        Match time = Regex.Match(
            match.Groups[2].Value,
            "^([0-9]{2}):([0-9]{2})(?::([0-9]{2})(?:\\.([0-9]+))?)?$",
            RegexOptions.CultureInvariant);
        string normalizedTime = time.Groups[1].Value + ":" + time.Groups[2].Value;
        string fraction = time.Groups[4].Success ? time.Groups[4].Value.TrimEnd('0') : string.Empty;
        if (time.Groups[3].Success && (time.Groups[3].Value != "00" || fraction.Length > 0)) {
            normalizedTime += ":" + time.Groups[3].Value;
            if (fraction.Length > 0) normalizedTime += "." + fraction;
        }
        return match.Groups[1].Value + "T" + normalizedTime;
    }

    private static bool TryReadPositiveYear(string value, out int yearModulo400) {
        yearModulo400 = 0;
        bool positive = false;
        foreach (char character in value) {
            if (character < '0' || character > '9') return false;
            int digit = character - '0';
            if (digit != 0) positive = true;
            yearModulo400 = (yearModulo400 * 10 + digit) % 400;
        }
        return positive;
    }

    private static int DaysInMonth(int yearModulo400, int month) {
        if (month < 1 || month > 12) return 0;
        if (month == 2) return IsLeapYear(yearModulo400) ? 29 : 28;
        return month == 4 || month == 6 || month == 9 || month == 11 ? 30 : 31;
    }

    private static bool IsLeapYear(int yearModulo400) =>
        yearModulo400 % 4 == 0 && (yearModulo400 % 100 != 0 || yearModulo400 == 0);

    private static int WeeksInYear(int yearModulo400) {
        int januaryFirst = DayOfWeek(yearModulo400, 1, 1);
        return januaryFirst == 4 || januaryFirst == 3 && IsLeapYear(yearModulo400) ? 53 : 52;
    }

    private static int DayOfWeek(int yearModulo400, int month, int day) {
        int cycleYear = yearModulo400 == 0 ? 400 : yearModulo400;
        int adjustedYear = month < 3 ? cycleYear - 1 : cycleYear;
        int adjustedMonth = month < 3 ? month + 12 : month;
        int value = adjustedYear + adjustedYear / 4 - adjustedYear / 100 + adjustedYear / 400
            + (153 * adjustedMonth - 457) / 5 + day - 306;
        return (value % 7 + 7) % 7;
    }

    private static string StripLineBreaks(string value) => value.Replace("\r", string.Empty).Replace("\n", string.Empty);

    private static string TrimAsciiWhitespace(string value) =>
        value.Trim(' ', '\t', '\r', '\n', '\f');

    private static double Midpoint(double minimum, double maximum) {
        double span = maximum - minimum;
        return double.IsInfinity(span)
            ? minimum / 2D + maximum / 2D
            : minimum + span / 2D;
    }

    private static double ResolveRangeFraction(double value, double minimum, double maximum) {
        if (maximum <= minimum) return 0D;
        double span = maximum - minimum;
        double offset = value - minimum;
        if (double.IsInfinity(span) || double.IsInfinity(offset)) {
            span = maximum / 2D - minimum / 2D;
            offset = value / 2D - minimum / 2D;
        }
        if (span <= 0D || double.IsNaN(span) || double.IsNaN(offset)) return 0D;
        return Math.Max(0D, Math.Min(1D, offset / span));
    }

    private static string FormatHtmlNumber(double value) =>
        value == 0D ? "0" : value.ToString("R", CultureInfo.InvariantCulture);

    private static bool IsDescendantOf(IElement element, IElement ancestor) {
        for (IElement? current = element; current != null; current = current.ParentElement) {
            if (ReferenceEquals(current, ancestor)) return true;
        }
        return false;
    }

    private static bool IsValidInputType(string type) {
        switch (type) {
            case "button":
            case "checkbox":
            case "color":
            case "date":
            case "datetime-local":
            case "email":
            case "file":
            case "hidden":
            case "image":
            case "month":
            case "number":
            case "password":
            case "radio":
            case "range":
            case "reset":
            case "search":
            case "submit":
            case "tel":
            case "text":
            case "time":
            case "url":
            case "week":
                return true;
            default:
                return false;
        }
    }

    private static bool IsValidButtonType(string type) =>
        type == "submit" || type == "reset" || type == "button";

    private static bool IsTextInputType(string type) {
        switch (type) {
            case "email":
            case "password":
            case "search":
            case "tel":
            case "text":
            case "url":
                return true;
            default:
                return false;
        }
    }

    private static string NormalizeIdentifier(string? value) => (value ?? string.Empty).Trim().ToLowerInvariant();

    private static string NormalizeKeyword(string? value) {
        string text = value ?? string.Empty;
        char[]? normalized = null;
        for (int index = 0; index < text.Length; index++) {
            char current = text[index];
            if (current < 'A' || current > 'Z') continue;
            normalized ??= text.ToCharArray();
            normalized[index] = (char)(current + ('a' - 'A'));
        }
        return normalized == null ? text : new string(normalized);
    }

    internal static string NormalizeOptionText(string? value) {
        string text = value ?? string.Empty;
        var normalized = new StringBuilder(text.Length);
        bool hasText = false;
        bool pendingSpace = false;
        foreach (char current in text) {
            if (IsAsciiWhitespace(current)) {
                if (hasText) pendingSpace = true;
                continue;
            }
            if (pendingSpace) {
                normalized.Append(' ');
                pendingSpace = false;
            }
            normalized.Append(current);
            hasText = true;
        }
        return normalized.ToString();
    }

    private static bool IsAsciiWhitespace(char value) =>
        value == '\t' || value == '\n' || value == '\f' || value == '\r' || value == ' ';

    private readonly struct HtmlRangeState {
        internal HtmlRangeState(string valueText, double fraction) {
            ValueText = valueText;
            Fraction = fraction;
        }

        internal string ValueText { get; }
        internal double Fraction { get; }
    }
}
