namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    /// <summary>Catalog optional-content/layer metadata discovered from /OCProperties.</summary>
    public PdfOptionalContentProperties? OptionalContent => ReadLogicalContent(_optionalContent);

    private PdfOptionalContentProperties? ExtractOptionalContent() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue("OCProperties", out PdfObject? optionalContentObject) ||
            ResolveObject(optionalContentObject) is not PdfDictionary optionalContent) {
            return null;
        }

        PdfArray? groupArray = ResolveArray(optionalContent.Items.TryGetValue("OCGs", out PdfObject? groupsObject) ? groupsObject : null);
        if (groupArray is null || groupArray.Items.Count == 0) {
            return new PdfOptionalContentProperties(
                Array.Empty<PdfOptionalContentGroup>(),
                null,
                null,
                null,
                Array.Empty<int>(),
                Array.Empty<int>(),
                Array.Empty<int>(),
                Array.Empty<int>());
        }

        PdfDictionary? defaultConfig = ResolveDict(optionalContent.Items.TryGetValue("D", out PdfObject? defaultConfigObject) ? defaultConfigObject : null);
        string? defaultName = defaultConfig is null ? null : TryReadText(defaultConfig, "Name");
        string? defaultCreator = defaultConfig is null ? null : TryReadText(defaultConfig, "Creator");
        string? baseState = defaultConfig is null ? null : TryReadName(defaultConfig, "BaseState");
        IReadOnlyList<int> onGroups = defaultConfig is null ? Array.Empty<int>() : ReadReferenceObjectNumbers(defaultConfig, "ON", includeNestedArrays: false);
        IReadOnlyList<int> offGroups = defaultConfig is null ? Array.Empty<int>() : ReadReferenceObjectNumbers(defaultConfig, "OFF", includeNestedArrays: false);
        IReadOnlyList<int> lockedGroups = defaultConfig is null ? Array.Empty<int>() : ReadReferenceObjectNumbers(defaultConfig, "Locked", includeNestedArrays: false);
        IReadOnlyList<int> orderGroups = defaultConfig is null ? Array.Empty<int>() : ReadReferenceObjectNumbers(defaultConfig, "Order", includeNestedArrays: true);
        var onSet = new HashSet<int>(onGroups);
        var offSet = new HashSet<int>(offGroups);
        var lockedSet = new HashSet<int>(lockedGroups);
        var orderSet = new HashSet<int>(orderGroups);

        var groups = new List<PdfOptionalContentGroup>();
        for (int i = 0; i < groupArray.Items.Count; i++) {
            PdfObject item = groupArray.Items[i];
            int? objectNumber = item is PdfReference reference ? reference.ObjectNumber : null;
            PdfDictionary? group = ResolveDict(item);
            if (group is null) {
                continue;
            }

            string? name = TryReadText(group, "Name");
            if (string.IsNullOrWhiteSpace(name)) {
                continue;
            }

            bool? isInitiallyVisible = GetInitialVisibility(objectNumber, baseState, onSet, offSet);
            PdfDictionary? usage = ResolveDict(group.Items.TryGetValue("Usage", out PdfObject? usageObject) ? usageObject : null);
            PdfDictionary? creatorInfo = ResolveDict(usage?.Items.TryGetValue("CreatorInfo", out PdfObject? creatorInfoObject) == true ? creatorInfoObject : null);
            PdfDictionary? view = ResolveDict(usage?.Items.TryGetValue("View", out PdfObject? viewObject) == true ? viewObject : null);
            PdfDictionary? print = ResolveDict(usage?.Items.TryGetValue("Print", out PdfObject? printObject) == true ? printObject : null);
            PdfDictionary? export = ResolveDict(usage?.Items.TryGetValue("Export", out PdfObject? exportObject) == true ? exportObject : null);

            groups.Add(new PdfOptionalContentGroup(
                objectNumber,
                name!,
                ReadNameList(group, "Intent"),
                isInitiallyVisible,
                objectNumber.HasValue && lockedSet.Contains(objectNumber.Value),
                objectNumber.HasValue && orderSet.Contains(objectNumber.Value),
                view is null ? null : TryReadName(view, "ViewState"),
                print is null ? null : TryReadName(print, "PrintState"),
                export is null ? null : TryReadName(export, "ExportState"),
                creatorInfo is null ? null : TryReadText(creatorInfo, "Creator"),
                creatorInfo is null ? null : TryReadName(creatorInfo, "Subtype")));
        }

        return new PdfOptionalContentProperties(
            groups.Count == 0 ? Array.Empty<PdfOptionalContentGroup>() : groups.AsReadOnly(),
            defaultName,
            defaultCreator,
            baseState,
            onGroups,
            offGroups,
            lockedGroups,
            orderGroups);
    }

    private static bool? GetInitialVisibility(int? objectNumber, string? baseState, HashSet<int> onSet, HashSet<int> offSet) {
        if (objectNumber.HasValue) {
            if (onSet.Contains(objectNumber.Value)) {
                return true;
            }

            if (offSet.Contains(objectNumber.Value)) {
                return false;
            }
        }

        switch (baseState) {
            case "ON":
                return true;
            case "OFF":
                return false;
            default:
                return null;
        }
    }

    private IReadOnlyList<string> ReadNameList(PdfDictionary dictionary, string key) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value)) {
            return Array.Empty<string>();
        }

        PdfObject? resolved = ResolveObject(value);
        if (resolved is PdfName name && !string.IsNullOrEmpty(name.Name)) {
            return new[] { name.Name };
        }

        if (resolved is not PdfArray array) {
            return Array.Empty<string>();
        }

        var names = new List<string>();
        for (int i = 0; i < array.Items.Count; i++) {
            PdfObject? item = ResolveObject(array.Items[i]);
            if (item is PdfName itemName && !string.IsNullOrEmpty(itemName.Name)) {
                names.Add(itemName.Name);
            } else if (item is PdfStringObj itemText && !string.IsNullOrEmpty(itemText.Value)) {
                names.Add(itemText.Value);
            }
        }

        return names.Count == 0 ? Array.Empty<string>() : names.AsReadOnly();
    }

    private IReadOnlyList<int> ReadReferenceObjectNumbers(PdfDictionary dictionary, string key, bool includeNestedArrays) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value) ||
            ResolveArray(value) is not PdfArray array) {
            return Array.Empty<int>();
        }

        var objectNumbers = new List<int>();
        AddReferenceObjectNumbers(array, objectNumbers, includeNestedArrays);
        return objectNumbers.Count == 0 ? Array.Empty<int>() : objectNumbers.AsReadOnly();
    }

    private void AddReferenceObjectNumbers(PdfArray array, List<int> objectNumbers, bool includeNestedArrays) {
        for (int i = 0; i < array.Items.Count; i++) {
            PdfObject item = array.Items[i];
            if (item is PdfReference reference) {
                if (!objectNumbers.Contains(reference.ObjectNumber)) {
                    objectNumbers.Add(reference.ObjectNumber);
                }

                continue;
            }

            if (includeNestedArrays && ResolveArray(item) is PdfArray nested) {
                AddReferenceObjectNumbers(nested, objectNumbers, includeNestedArrays);
            }
        }
    }
}
