namespace OfficeIMO.Pdf;

internal sealed class PdfPageOptionalContentVisibility {
    private readonly Dictionary<string, bool> _hiddenProperties;
    private readonly HashSet<int> _hiddenObjectNumbers;

    private PdfPageOptionalContentVisibility(Dictionary<string, bool> hiddenProperties, HashSet<int> hiddenObjectNumbers) {
        _hiddenProperties = hiddenProperties;
        _hiddenObjectNumbers = hiddenObjectNumbers;
    }

    public static PdfPageOptionalContentVisibility? Create(
        PdfDictionary? resources,
        Dictionary<int, PdfIndirectObject> objects) {
        Dictionary<int, bool> groupVisibility = ReadGroupVisibility(objects);
        if (groupVisibility.Count == 0) {
            return null;
        }

        var hiddenObjectNumbers = new HashSet<int>();
        foreach (KeyValuePair<int, bool> entry in groupVisibility) {
            if (!entry.Value) {
                hiddenObjectNumbers.Add(entry.Key);
            }
        }

        var hiddenProperties = new Dictionary<string, bool>(StringComparer.Ordinal);
        if (resources != null &&
            resources.Items.TryGetValue("Properties", out PdfObject? propertiesObject) &&
            ResolveObject(propertiesObject, objects) is PdfDictionary properties) {
            foreach (KeyValuePair<string, PdfObject> entry in properties.Items) {
                if (TryGetReferencedObjectNumber(entry.Value, out int objectNumber) &&
                    groupVisibility.TryGetValue(objectNumber, out bool isVisible) &&
                    !isVisible) {
                    hiddenProperties[entry.Key] = true;
                }
            }
        }

        return hiddenProperties.Count == 0 && hiddenObjectNumbers.Count == 0
            ? null
            : new PdfPageOptionalContentVisibility(hiddenProperties, hiddenObjectNumbers);
    }

    public bool IsHidden(string propertyName) =>
        _hiddenProperties.TryGetValue(propertyName, out bool hidden) && hidden;

    public bool IsHiddenAny(IReadOnlyList<int> objectNumbers) {
        for (int i = 0; i < objectNumbers.Count; i++) {
            if (_hiddenObjectNumbers.Contains(objectNumbers[i])) {
                return true;
            }
        }

        return false;
    }

    private static Dictionary<int, bool> ReadGroupVisibility(Dictionary<int, PdfIndirectObject> objects) {
        var result = new Dictionary<int, bool>();
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects);
        if (catalog == null ||
            !catalog.Items.TryGetValue("OCProperties", out PdfObject? optionalContentObject) ||
            ResolveObject(optionalContentObject, objects) is not PdfDictionary optionalContent ||
            ResolveObject(optionalContent.Items.TryGetValue("OCGs", out PdfObject? groupsObject) ? groupsObject : null, objects) is not PdfArray groups) {
            return result;
        }

        PdfDictionary? defaultConfiguration = ResolveObject(
            optionalContent.Items.TryGetValue("D", out PdfObject? defaultConfigurationObject) ? defaultConfigurationObject : null,
            objects) as PdfDictionary;
        string? baseState = ReadName(defaultConfiguration, "BaseState", objects);
        HashSet<int> onGroups = ReadReferenceSet(defaultConfiguration, "ON", objects);
        HashSet<int> offGroups = ReadReferenceSet(defaultConfiguration, "OFF", objects);

        for (int i = 0; i < groups.Items.Count; i++) {
            if (groups.Items[i] is not PdfReference reference) {
                continue;
            }

            bool isVisible = true;
            if (string.Equals(baseState, "OFF", StringComparison.Ordinal)) {
                isVisible = onGroups.Contains(reference.ObjectNumber);
            } else if (offGroups.Contains(reference.ObjectNumber)) {
                isVisible = false;
            } else if (onGroups.Contains(reference.ObjectNumber)) {
                isVisible = true;
            }

            result[reference.ObjectNumber] = isVisible;
        }

        return result;
    }

    private static HashSet<int> ReadReferenceSet(PdfDictionary? dictionary, string key, Dictionary<int, PdfIndirectObject> objects) {
        var result = new HashSet<int>();
        if (dictionary == null ||
            ResolveObject(dictionary.Items.TryGetValue(key, out PdfObject? value) ? value : null, objects) is not PdfArray array) {
            return result;
        }

        for (int i = 0; i < array.Items.Count; i++) {
            AddReferenceObjectNumbers(array.Items[i], objects, result);
        }

        return result;
    }

    private static void AddReferenceObjectNumbers(PdfObject value, Dictionary<int, PdfIndirectObject> objects, HashSet<int> result) {
        if (value is PdfReference reference) {
            result.Add(reference.ObjectNumber);
            return;
        }

        if (ResolveObject(value, objects) is PdfArray nested) {
            for (int i = 0; i < nested.Items.Count; i++) {
                AddReferenceObjectNumbers(nested.Items[i], objects, result);
            }
        }
    }

    private static string? ReadName(PdfDictionary? dictionary, string key, Dictionary<int, PdfIndirectObject> objects) {
        if (dictionary == null ||
            ResolveObject(dictionary.Items.TryGetValue(key, out PdfObject? value) ? value : null, objects) is not PdfName name ||
            string.IsNullOrEmpty(name.Name)) {
            return null;
        }

        return name.Name;
    }

    private static bool TryGetReferencedObjectNumber(PdfObject value, out int objectNumber) {
        if (value is PdfReference reference) {
            objectNumber = reference.ObjectNumber;
            return true;
        }

        objectNumber = 0;
        return false;
    }

    private static PdfObject? ResolveObject(PdfObject? value, Dictionary<int, PdfIndirectObject> objects) =>
        PdfObjectLookup.Resolve(objects, value);
}
