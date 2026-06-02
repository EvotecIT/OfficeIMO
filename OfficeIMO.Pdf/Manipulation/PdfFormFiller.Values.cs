namespace OfficeIMO.Pdf;

public static partial class PdfFormFiller {
    private readonly struct ChoiceFillValue {
        public string ExportValue { get; }
        public string DisplayValue { get; }

        public ChoiceFillValue(string exportValue, string displayValue) {
            ExportValue = exportValue;
            DisplayValue = displayValue;
        }
    }

    private static string? TryReadSimpleValue(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out var value)) {
            return null;
        }

        return ResolveObject(objects, value) switch {
            PdfStringObj text => text.Value,
            PdfName name => name.Name,
            PdfNumber number => FormatNumber(number.Value),
            PdfBoolean boolean => boolean.Value ? "true" : "false",
            _ => null
        };
    }

    private static IReadOnlyList<string>? TryReadSimpleValues(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out var value)) {
            return null;
        }

        PdfObject? resolved = ResolveObject(objects, value);
        if (resolved is PdfArray array) {
            var values = new List<string>();
            for (int i = 0; i < array.Items.Count; i++) {
                if (TryFormatSimpleValue(objects, array.Items[i], out string? item)) {
                    values.Add(item!);
                }
            }

            return values.Count == 0 ? null : values.AsReadOnly();
        }

        return TryFormatSimpleValue(objects, resolved, out string? single)
            ? new[] { single! }
            : null;
    }

    private static bool TryFormatSimpleValue(Dictionary<int, PdfIndirectObject> objects, PdfObject? value, out string? text) {
        text = null;
        switch (ResolveObject(objects, value)) {
            case PdfStringObj stringObj:
                text = stringObj.Value;
                return true;
            case PdfName name:
                text = name.Name;
                return true;
            case PdfNumber number:
                text = FormatNumber(number.Value);
                return true;
            case PdfBoolean boolean:
                text = boolean.Value ? "true" : "false";
                return true;
            default:
                return false;
        }
    }

    private static string? JoinSimpleValues(IReadOnlyList<string>? values) {
        return values is { Count: > 0 }
            ? string.Join(", ", values)
            : null;
    }

    private static string? TryResolveChoiceDisplayValue(Dictionary<int, PdfIndirectObject> objects, PdfArray? options, IReadOnlyList<string>? exportValues) {
        if (exportValues is not { Count: > 0 } ||
            options is null ||
            options.Items.Count == 0) {
            return null;
        }

        var displayValues = new List<string>(exportValues.Count);
        for (int i = 0; i < exportValues.Count; i++) {
            displayValues.Add(ResolveSingleChoiceDisplayValue(objects, options, exportValues[i]) ?? exportValues[i]);
        }

        return displayValues.Count == 0 ? null : string.Join(", ", displayValues);
    }

    private static string? ResolveSingleChoiceDisplayValue(Dictionary<int, PdfIndirectObject> objects, PdfArray options, string exportValue) {
        for (int i = 0; i < options.Items.Count; i++) {
            PdfObject? optionObject = ResolveObject(objects, options.Items[i]);
            if (optionObject is PdfArray pair &&
                pair.Items.Count >= 2 &&
                TryReadOptionText(objects, pair.Items[0], out string? pairExportValue) &&
                string.Equals(pairExportValue, exportValue, StringComparison.Ordinal) &&
                TryReadOptionText(objects, pair.Items[1], out string? pairDisplayText)) {
                return pairDisplayText;
            }

            if (optionObject is not null &&
                TryReadOptionText(objects, optionObject, out string? singleValue) &&
                string.Equals(singleValue, exportValue, StringComparison.Ordinal)) {
                return singleValue;
            }
        }

        return null;
    }

    private static List<ChoiceFillValue> ResolveChoiceFillValues(Dictionary<int, PdfIndirectObject> objects, PdfArray? options, bool isEditableChoice, IReadOnlyList<string> values) {
        var resolved = new List<ChoiceFillValue>(values.Count);
        for (int i = 0; i < values.Count; i++) {
            resolved.Add(options is null || options.Items.Count == 0
                ? new ChoiceFillValue(values[i], values[i])
                : ResolveSingleChoiceFillValue(objects, options, isEditableChoice, values[i]));
        }

        return resolved;
    }

    private static ChoiceFillValue ResolveSingleChoiceFillValue(Dictionary<int, PdfIndirectObject> objects, PdfArray options, bool isEditableChoice, string value) {
        for (int i = 0; i < options.Items.Count; i++) {
            PdfObject? optionObject = ResolveObject(objects, options.Items[i]);
            if (optionObject is PdfArray pair &&
                pair.Items.Count >= 2 &&
                TryReadOptionText(objects, pair.Items[0], out string? pairExportValue) &&
                pairExportValue is not null &&
                TryReadOptionText(objects, pair.Items[1], out string? pairDisplayText) &&
                pairDisplayText is not null &&
                string.Equals(pairExportValue, value, StringComparison.Ordinal)) {
                return new ChoiceFillValue(pairExportValue, pairDisplayText);
            }

            if (optionObject is not null &&
                TryReadOptionText(objects, optionObject, out string? singleValue) &&
                singleValue is not null &&
                string.Equals(singleValue, value, StringComparison.Ordinal)) {
                return new ChoiceFillValue(singleValue, singleValue);
            }
        }

        for (int i = 0; i < options.Items.Count; i++) {
            PdfObject? optionObject = ResolveObject(objects, options.Items[i]);
            if (optionObject is PdfArray pair &&
                pair.Items.Count >= 2 &&
                TryReadOptionText(objects, pair.Items[0], out string? pairExportValue) &&
                pairExportValue is not null &&
                TryReadOptionText(objects, pair.Items[1], out string? pairDisplayText) &&
                pairDisplayText is not null &&
                string.Equals(pairDisplayText, value, StringComparison.Ordinal)) {
                return new ChoiceFillValue(pairExportValue, pairDisplayText);
            }
        }

        if (isEditableChoice) {
            return new ChoiceFillValue(value, value);
        }

        throw new ArgumentException("PDF choice field value does not match an available option: " + value, nameof(value));
    }

    private static bool TryReadOptionText(Dictionary<int, PdfIndirectObject> objects, PdfObject value, out string? text) {
        switch (ResolveObject(objects, value)) {
            case PdfStringObj stringObj:
                text = stringObj.Value;
                return true;
            case PdfName name:
                text = name.Name;
                return true;
            case PdfNumber number:
                text = FormatNumber(number.Value);
                return true;
            default:
                text = null;
                return false;
        }
    }
}
