using AngleSharp.Dom;
using System.Globalization;
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
        switch (NormalizeIdentifier(attributeName)) {
            case "checked": return IsCheckedStateApplicable(elementName, effectiveType);
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
            case "value": return NormalizeIdentifier(elementName) != "input" || effectiveType != "file";
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
                    element.GetAttribute("max")) };
            }
        }

        string? authored = element.GetAttribute("value");
        if (authored != null) return new[] { authored };
        string defaultValue = GetDefaultValue(name, element.GetAttribute("type"), element.TextContent ?? string.Empty);
        return defaultValue.Length == 0 ? Array.Empty<string>() : new[] { defaultValue };
    }

    internal static string GetDefaultValue(string elementName, string? type, string textContent) {
        string name = NormalizeIdentifier(elementName);
        if (name == "option") return NormalizeText(textContent);
        if (name == "input") {
            string effectiveType = GetEffectiveType(name, type);
            if (effectiveType == "checkbox" || effectiveType == "radio") return "on";
        }
        return string.Empty;
    }

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
        return (owner?.GetAttribute("id") ?? string.Empty).Trim();
    }

    internal static IElement? ResolveFormOwner(IElement element) {
        if (element.HasAttribute("form")) {
            string explicitOwner = (element.GetAttribute("form") ?? string.Empty).Trim();
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

    private static IReadOnlyList<string> GetSelectValues(IElement select) {
        bool multiple = select.HasAttribute("multiple");
        IElement[] options = select.QuerySelectorAll("option").ToArray();
        IElement[] selected = options.Where(option => option.HasAttribute("selected")).ToArray();
        if (multiple) return selected.Select(GetOptionValue).ToArray();
        IElement? effective = selected.LastOrDefault() ?? options.FirstOrDefault();
        return effective == null ? Array.Empty<string>() : new[] { GetOptionValue(effective) };
    }

    private static string GetOptionValue(IElement option) =>
        option.GetAttribute("value") ?? GetDefaultValue("option", null, option.TextContent ?? string.Empty);

    internal static string GetRangeValue(string? value, string? minimum, string? maximum) =>
        ResolveRange(value, minimum, maximum).ValueText;

    internal static double GetRangeFraction(IElement element) => ResolveRange(
        element.GetAttribute("value"),
        element.GetAttribute("min"),
        element.GetAttribute("max")).Fraction;

    private static HtmlRangeState ResolveRange(string? value, string? minimum, string? maximum) {
        double min = TryParseHtmlNumber(minimum, out double parsedMinimum) ? parsedMinimum : 0D;
        double max = TryParseHtmlNumber(maximum, out double parsedMaximum) ? parsedMaximum : 100D;
        if (max < min) return new HtmlRangeState(FormatHtmlNumber(min), 0D);

        bool preservesAuthoredValue = TryParseHtmlNumber(value, out double current);
        if (!preservesAuthoredValue) current = Midpoint(min, max);
        if (current < min) {
            current = min;
            preservesAuthoredValue = false;
        } else if (current > max) {
            current = max;
            preservesAuthoredValue = false;
        }

        double fraction = ResolveRangeFraction(current, min, max);
        return new HtmlRangeState(
            preservesAuthoredValue ? value! : FormatHtmlNumber(current),
            fraction);
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

    private static string NormalizeText(string? value) =>
        string.Join(" ", (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private readonly struct HtmlRangeState {
        internal HtmlRangeState(string valueText, double fraction) {
            ValueText = valueText;
            Fraction = fraction;
        }

        internal string ValueText { get; }
        internal double Fraction { get; }
    }
}
