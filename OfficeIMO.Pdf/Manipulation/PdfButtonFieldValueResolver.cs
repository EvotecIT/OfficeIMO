namespace OfficeIMO.Pdf;

/// <summary>Maps external button values to the appearance-state names stored by AcroForm widgets.</summary>
internal static class PdfButtonFieldValueResolver {
    internal static string Resolve(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary field,
        PdfArray? inheritedOptions,
        IReadOnlyCollection<string> availableStates,
        bool isRadioButtonGroup,
        string value) {
        if (TryResolveExportAppearanceState(objects, field, inheritedOptions, availableStates, isRadioButtonGroup, value, out string? appearanceState)) {
            return appearanceState!;
        }
        if (availableStates.Contains(value, StringComparer.Ordinal)) return value;
        if (IsOffValue(value)) return "Off";
        if (!isRadioButtonGroup && IsTruthyValue(value)) {
            if (availableStates.Count == 0) return "Yes";
            if (availableStates.Count == 1) return availableStates.Single();
        }

        string fieldKind = isRadioButtonGroup ? "radio button" : "checkbox";
        throw new ArgumentException($"PDF {fieldKind} field cannot be filled with value '{value}' because it is not one of the available appearance states.", nameof(value));
    }

    private static bool TryResolveExportAppearanceState(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary field,
        PdfArray? inheritedOptions,
        IReadOnlyCollection<string> availableStates,
        bool isRadioButtonGroup,
        string value,
        out string? appearanceState) {
        appearanceState = null;
        PdfArray? options = field.Items.TryGetValue("Opt", out PdfObject? optionsObject)
            ? ResolveObject(objects, optionsObject) as PdfArray
            : inheritedOptions;
        if (options is null) return false;
        var exportValues = new List<string>();
        foreach (PdfObject option in options.Items) {
            if (!TryReadOptionText(objects, option, out string? exportValue) || exportValue is null) return false;
            exportValues.Add(exportValue);
        }
        var appearanceStates = new List<string>();
        CollectOrderedOnStates(objects, field, appearanceStates, new HashSet<int>());
        int[] matches = exportValues
            .Select((exportValue, index) => new { exportValue, index })
            .Where(item => string.Equals(item.exportValue, value, StringComparison.Ordinal))
            .Select(item => item.index)
            .ToArray();
        if (matches.Length == 0) return false;
        if (matches.Length > 1) {
            throw new ArgumentException($"PDF button field export value '{value}' is ambiguous because it occurs more than once.", nameof(value));
        }
        int matchedIndex = matches[0];
        if (!isRadioButtonGroup && exportValues.Count == 1) {
            if (appearanceStates.Count == 0) {
                appearanceState = "Yes";
                return true;
            }
            if (appearanceStates.Count == 1) {
                appearanceState = appearanceStates[0];
                return true;
            }
        }
        if (exportValues.Count != appearanceStates.Count || matchedIndex >= appearanceStates.Count) return false;
        appearanceState = appearanceStates[matchedIndex];
        return availableStates.Contains(appearanceState, StringComparer.Ordinal);
    }

    private static void CollectOrderedOnStates(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary field,
        ICollection<string> states,
        ISet<int> visited) {
        if (field.Items.TryGetValue("AP", out PdfObject? appearanceObject)
            && ResolveObject(objects, appearanceObject) is PdfDictionary appearance
            && appearance.Items.TryGetValue("N", out PdfObject? normalObject)
            && ResolveObject(objects, normalObject) is PdfDictionary normal) {
            foreach (string state in normal.Items.Keys) {
                if (!string.Equals(state, "Off", StringComparison.Ordinal) && !states.Contains(state)) states.Add(state);
            }
        }
        if (!field.Items.TryGetValue("Kids", out PdfObject? kidsObject)
            || ResolveObject(objects, kidsObject) is not PdfArray kids) {
            return;
        }
        foreach (PdfObject kidObject in kids.Items) {
            if (kidObject is PdfReference reference && !visited.Add(reference.ObjectNumber)) continue;
            if (ResolveObject(objects, kidObject) is PdfDictionary kid) {
                CollectOrderedOnStates(objects, kid, states, visited);
            }
        }
    }

    private static bool TryReadOptionText(Dictionary<int, PdfIndirectObject> objects, PdfObject value, out string? text) {
        PdfObject? resolved = ResolveObject(objects, value);
        if (resolved is PdfArray pair && pair.Items.Count > 0) resolved = ResolveObject(objects, pair.Items[0]);
        switch (resolved) {
            case PdfStringObj stringValue:
                text = stringValue.Value;
                return true;
            case PdfName name:
                text = name.Name;
                return true;
            default:
                text = null;
                return false;
        }
    }

    private static PdfObject? ResolveObject(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) {
        var visited = new HashSet<int>();
        while (value is PdfReference reference && visited.Add(reference.ObjectNumber)) {
            if (!objects.TryGetValue(reference.ObjectNumber, out PdfIndirectObject? indirect)) return null;
            value = indirect.Value;
        }
        return value;
    }

    private static bool IsOffValue(string value) =>
        string.IsNullOrWhiteSpace(value)
        || string.Equals(value, "false", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "off", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "0", StringComparison.Ordinal);

    private static bool IsTruthyValue(string value) =>
        string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "on", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "1", StringComparison.Ordinal);
}
