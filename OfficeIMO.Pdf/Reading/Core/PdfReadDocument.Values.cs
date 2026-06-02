namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private static string? CombineFieldName(string? parentName, string? partialName) {
        if (string.IsNullOrEmpty(parentName)) {
            return string.IsNullOrEmpty(partialName) ? null : partialName;
        }

        if (string.IsNullOrEmpty(partialName)) {
            return parentName;
        }

        return parentName + "." + partialName;
    }

    private string? TryReadText(PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out var value) && ResolveObject(value) is PdfStringObj text && !string.IsNullOrEmpty(text.Value)
            ? text.Value
            : null;
    }

    private string? TryReadName(PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out var value) && ResolveObject(value) is PdfName name && !string.IsNullOrEmpty(name.Name)
            ? name.Name
            : null;
    }

    private string? TryReadSimpleFieldValue(PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out var value) || !TryFormatSimpleValue(value, out string? text)) {
            return null;
        }

        return text;
    }

    private IReadOnlyList<string> ReadSimpleFieldValues(PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out var value)) {
            return Array.Empty<string>();
        }

        PdfObject? resolved = ResolveObject(value);
        if (resolved is PdfArray array) {
            var values = new List<string>();
            for (int i = 0; i < array.Items.Count; i++) {
                if (TryFormatSimpleValue(array.Items[i], out string? itemText)) {
                    values.Add(itemText!);
                }
            }

            return values.Count == 0 ? Array.Empty<string>() : values.AsReadOnly();
        }

        if (resolved is not null && TryFormatSimpleValue(resolved, out string? text)) {
            return new[] { text! };
        }

        return Array.Empty<string>();
    }

    private IReadOnlyList<PdfFormFieldOption> ReadFormFieldOptions(PdfDictionary dictionary) {
        if (!dictionary.Items.TryGetValue("Opt", out var optionsObject) ||
            ResolveArray(optionsObject) is not PdfArray optionsArray ||
            optionsArray.Items.Count == 0) {
            return Array.Empty<PdfFormFieldOption>();
        }

        var options = new List<PdfFormFieldOption>();
        for (int i = 0; i < optionsArray.Items.Count; i++) {
            PdfObject? optionObject = ResolveObject(optionsArray.Items[i]);
            if (optionObject is PdfArray pair &&
                pair.Items.Count >= 2 &&
                TryReadOptionText(pair.Items[0], out string? exportValue) &&
                TryReadOptionText(pair.Items[1], out string? displayText)) {
                options.Add(new PdfFormFieldOption(exportValue!, displayText!));
                continue;
            }

            if (optionObject is not null && TryReadOptionText(optionObject, out string? value)) {
                options.Add(new PdfFormFieldOption(value!, value!));
            }
        }

        return options.Count == 0 ? Array.Empty<PdfFormFieldOption>() : options.AsReadOnly();
    }

    private bool TryReadOptionText(PdfObject value, out string? text) {
        PdfObject? resolved = ResolveObject(value);
        switch (resolved) {
            case PdfStringObj stringObj:
                text = stringObj.Value;
                return true;
            case PdfName name:
                text = name.Name;
                return true;
            case PdfNumber number:
                text = number.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                return true;
            default:
                text = null;
                return false;
        }
    }

    private int? TryReadInteger(PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out var value) ||
            ResolveObject(value) is not PdfNumber number ||
            number.Value < int.MinValue ||
            number.Value > int.MaxValue ||
            Math.Truncate(number.Value) != number.Value) {
            return null;
        }

        return (int)number.Value;
    }

    private int? TryReadPositiveInteger(PdfDictionary dictionary, string key) {
        int? value = TryReadInteger(dictionary, key);
        return value.HasValue && value.Value > 0 ? value.Value : null;
    }

    private static bool TryGetNonNegativeInteger(PdfNumber number, out int value) {
        value = 0;
        if (number.Value < 0 || number.Value > int.MaxValue || Math.Truncate(number.Value) != number.Value) {
            return false;
        }

        value = (int)number.Value;
        return true;
    }

    private static bool TryGetPositiveInteger(PdfNumber number, out int value) {
        if (TryGetNonNegativeInteger(number, out value) && value > 0) {
            return true;
        }

        value = 0;
        return false;
    }

    private bool TryFormatSimpleValue(PdfObject value, out string? text) {
        switch (ResolveObject(value)) {
            case PdfNumber number:
                text = number.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                return true;
            case PdfBoolean boolean:
                text = boolean.Value ? "true" : "false";
                return true;
            case PdfName name:
                text = name.Name;
                return true;
            case PdfStringObj stringObj:
                text = stringObj.Value;
                return true;
            case PdfNull:
                text = "null";
                return true;
            case PdfArray array:
                var parts = new List<string>(array.Items.Count);
                for (int i = 0; i < array.Items.Count; i++) {
                    if (!TryFormatSimpleValue(array.Items[i], out string? itemText)) {
                        text = null;
                        return false;
                    }

                    parts.Add(itemText!);
                }

                text = "[" + string.Join(" ", parts) + "]";
                return true;
            default:
                text = null;
                return false;
        }
    }
}
