using AngleSharp.Dom;

namespace OfficeIMO.Html;

/// <summary>Canonical HTML form-control normalization shared by semantic projection and fidelity scoring.</summary>
internal static class HtmlFormControlSemantics {
    internal static string GetEffectiveType(string elementName, string? type, bool multiple = false) {
        string name = NormalizeName(elementName);
        string normalized = NormalizeName(type);
        if (name == "input") return IsValidInputType(normalized) ? normalized : "text";
        if (name == "button") return IsValidButtonType(normalized) ? normalized : "submit";
        if (name == "select") return multiple ? "select-multiple" : "select-one";
        return name;
    }

    internal static bool IsValidType(string elementName, string? type) {
        string name = NormalizeName(elementName);
        string normalized = NormalizeName(type);
        return name == "input" ? IsValidInputType(normalized)
            : name == "button" && IsValidButtonType(normalized);
    }

    internal static string GetEffectiveFormMethod(string? method) {
        string normalized = NormalizeName(method);
        return normalized == "post" || normalized == "dialog" ? normalized : "get";
    }

    internal static string GetEffectiveFormEncoding(string? encodingType) {
        string normalized = NormalizeName(encodingType);
        return normalized == "multipart/form-data" || normalized == "text/plain"
            ? normalized
            : "application/x-www-form-urlencoded";
    }

    internal static bool IsSubmitter(IElement element) {
        string name = NormalizeName(element.LocalName);
        string type = GetEffectiveType(name, element.GetAttribute("type"));
        return name == "button" ? type != "button" && type != "reset"
            : name == "input" && (type == "submit" || type == "image");
    }

    internal static bool IsCheckedStateApplicable(string elementName, string effectiveType) =>
        NormalizeName(elementName) == "input" && (effectiveType == "checkbox" || effectiveType == "radio");

    internal static bool IsMultipleStateApplicable(string elementName, string effectiveType) {
        string name = NormalizeName(elementName);
        return name == "select" || name == "input" && (effectiveType == "email" || effectiveType == "file");
    }

    internal static bool IsRequiredStateApplicable(string elementName, string effectiveType) {
        string name = NormalizeName(elementName);
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
        string name = NormalizeName(elementName);
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
        NormalizeName(elementName) == "input" && IsTextInputType(effectiveType);

    internal static bool IsLengthApplicable(string elementName, string effectiveType) =>
        NormalizeName(elementName) == "textarea"
        || NormalizeName(elementName) == "input" && IsTextInputType(effectiveType);

    internal static bool IsRangeApplicable(string elementName, string effectiveType) {
        if (NormalizeName(elementName) != "input") return false;
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
        NormalizeName(elementName) == "textarea"
        || NormalizeName(elementName) == "input" && IsTextInputType(effectiveType);

    internal static bool IsStateAttributeApplicable(string elementName, string effectiveType, string attributeName) {
        switch (NormalizeName(attributeName)) {
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
            case "value": return NormalizeName(elementName) != "input" || effectiveType != "file";
            default: return true;
        }
    }

    internal static IReadOnlyList<string> GetValues(IElement element) {
        string name = NormalizeName(element.LocalName);
        if (name == "select") return GetSelectValues(element);
        if (name == "textarea") return new[] { element.TextContent ?? string.Empty };
        if (name == "input" && GetEffectiveType(name, element.GetAttribute("type")) == "file") {
            return Array.Empty<string>();
        }

        string? authored = element.GetAttribute("value");
        if (authored != null) return new[] { authored };
        string defaultValue = GetDefaultValue(name, element.GetAttribute("type"), element.TextContent ?? string.Empty);
        return defaultValue.Length == 0 ? Array.Empty<string>() : new[] { defaultValue };
    }

    internal static string GetDefaultValue(string elementName, string? type, string textContent) {
        string name = NormalizeName(elementName);
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

    private static string NormalizeName(string? value) => (value ?? string.Empty).Trim().ToLowerInvariant();

    private static string NormalizeText(string? value) =>
        string.Join(" ", (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
}
