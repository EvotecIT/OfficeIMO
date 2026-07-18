namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private readonly struct ChoiceFillValue {
        public string ExportValue { get; }
        public string DisplayValue { get; }

        public ChoiceFillValue(string exportValue, string displayValue) {
            ExportValue = exportValue;
            DisplayValue = displayValue;
        }
    }

    private sealed class ChoiceOptionLookup {
        private readonly Dictionary<string, ChoiceFillValue> _byExportValue;
        private readonly Dictionary<string, ChoiceFillValue> _byDisplayValue;
        private readonly Dictionary<string, string> _displayByExportValue;

        private ChoiceOptionLookup(
            Dictionary<string, ChoiceFillValue> byExportValue,
            Dictionary<string, ChoiceFillValue> byDisplayValue,
            Dictionary<string, string> displayByExportValue) {
            _byExportValue = byExportValue;
            _byDisplayValue = byDisplayValue;
            _displayByExportValue = displayByExportValue;
        }

        public static ChoiceOptionLookup Create(Dictionary<int, PdfIndirectObject> objects, PdfArray options) {
            var byExportValue = new Dictionary<string, ChoiceFillValue>(StringComparer.Ordinal);
            var byDisplayValue = new Dictionary<string, ChoiceFillValue>(StringComparer.Ordinal);
            var displayByExportValue = new Dictionary<string, string>(StringComparer.Ordinal);

            for (int i = 0; i < options.Items.Count; i++) {
                PdfObject? optionObject = ResolveObject(objects, options.Items[i]);
                if (optionObject is PdfArray pair &&
                    pair.Items.Count >= 2 &&
                    TryReadOptionText(objects, pair.Items[0], out string? pairExportValue) &&
                    pairExportValue is not null &&
                    TryReadOptionText(objects, pair.Items[1], out string? pairDisplayText) &&
                    pairDisplayText is not null) {
                    var choice = new ChoiceFillValue(pairExportValue, pairDisplayText);
                    AddIfMissing(byExportValue, pairExportValue, choice);
                    AddIfMissing(displayByExportValue, pairExportValue, pairDisplayText);
                    AddIfMissing(byDisplayValue, pairDisplayText, choice);
                    continue;
                }

                if (optionObject is not null &&
                    TryReadOptionText(objects, optionObject, out string? singleValue) &&
                    singleValue is not null) {
                    var choice = new ChoiceFillValue(singleValue, singleValue);
                    AddIfMissing(byExportValue, singleValue, choice);
                    AddIfMissing(displayByExportValue, singleValue, singleValue);
                }
            }

            return new ChoiceOptionLookup(byExportValue, byDisplayValue, displayByExportValue);
        }

        public bool TryResolveFillValue(string value, out ChoiceFillValue fillValue) {
            if (_byExportValue.TryGetValue(value, out fillValue)) {
                return true;
            }

            return _byDisplayValue.TryGetValue(value, out fillValue);
        }

        public string? ResolveDisplayValue(string exportValue) =>
            _displayByExportValue.TryGetValue(exportValue, out string? displayValue) ? displayValue : null;

        private static void AddIfMissing(Dictionary<string, ChoiceFillValue> dictionary, string key, ChoiceFillValue value) {
#pragma warning disable CA1864 // Dictionary.TryAdd is unavailable for the net472 target.
            if (!dictionary.ContainsKey(key)) {
                dictionary.Add(key, value);
            }
#pragma warning restore CA1864
        }

        private static void AddIfMissing(Dictionary<string, string> dictionary, string key, string value) {
#pragma warning disable CA1864 // Dictionary.TryAdd is unavailable for the net472 target.
            if (!dictionary.ContainsKey(key)) {
                dictionary.Add(key, value);
            }
#pragma warning restore CA1864
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

    private static string? JoinSimpleValues(IReadOnlyList<string>? values, string separator) {
        return values is { Count: > 0 }
            ? string.Join(separator, values)
            : null;
    }

    private static string? TryResolveChoiceDisplayValue(Dictionary<int, PdfIndirectObject> objects, PdfArray? options, IReadOnlyList<string>? exportValues, string separator) {
        if (exportValues is not { Count: > 0 } ||
            options is null ||
            options.Items.Count == 0) {
            return null;
        }

        var displayValues = new List<string>(exportValues.Count);
        ChoiceOptionLookup optionLookup = ChoiceOptionLookup.Create(objects, options);
        for (int i = 0; i < exportValues.Count; i++) {
            displayValues.Add(optionLookup.ResolveDisplayValue(exportValues[i]) ?? exportValues[i]);
        }

        return displayValues.Count == 0 ? null : string.Join(separator, displayValues);
    }

    private static List<ChoiceFillValue> ResolveChoiceFillValues(Dictionary<int, PdfIndirectObject> objects, PdfArray? options, bool isEditableChoice, IReadOnlyList<string> values) {
        var resolved = new List<ChoiceFillValue>(values.Count);
        ChoiceOptionLookup? optionLookup = options is null || options.Items.Count == 0
            ? null
            : ChoiceOptionLookup.Create(objects, options);

        for (int i = 0; i < values.Count; i++) {
            resolved.Add(optionLookup is null
                ? new ChoiceFillValue(values[i], values[i])
                : ResolveChoiceFillValue(optionLookup, isEditableChoice, values[i]));
        }

        return resolved;
    }

    private static ChoiceFillValue ResolveChoiceFillValue(ChoiceOptionLookup optionLookup, bool isEditableChoice, string value) {
        if (optionLookup.TryResolveFillValue(value, out ChoiceFillValue fillValue)) {
            return fillValue;
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
