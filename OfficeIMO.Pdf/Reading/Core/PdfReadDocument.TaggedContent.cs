namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    /// <summary>Tagged PDF structure metadata discovered from /MarkInfo and /StructTreeRoot.</summary>
    public PdfTaggedContentInfo? TaggedContent => ReadLogicalContent(_taggedContent);

    /// <summary>True when a readable tagged-PDF structure tree was discovered.</summary>
    public bool HasTaggedContent => TaggedContent is not null;

    private PdfTaggedContentInfo? ExtractTaggedContent() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null) {
            return null;
        }

        PdfDictionary? markInfo = ResolveDict(catalog.Items.TryGetValue("MarkInfo", out PdfObject? markInfoObject) ? markInfoObject : null);
        PdfObject? structTreeRootObject = catalog.Items.TryGetValue("StructTreeRoot", out PdfObject? rootObject) ? rootObject : null;
        PdfDictionary? structTreeRoot = ResolveDict(structTreeRootObject);
        if (structTreeRoot is null) {
            return null;
        }

        int? structTreeRootObjectNumber = structTreeRootObject is PdfReference rootReference ? rootReference.ObjectNumber : null;
        int? parentTreeObjectNumber = null;
        PdfDictionary? parentTree = null;
        if (structTreeRoot is not null &&
            structTreeRoot.Items.TryGetValue("ParentTree", out PdfObject? parentTreeObject)) {
            if (parentTreeObject is PdfReference parentTreeReference) {
                parentTreeObjectNumber = parentTreeReference.ObjectNumber;
            }

            parentTree = ResolveDict(parentTreeObject);
        }

        return new PdfTaggedContentInfo(
            markInfo is null ? null : TryReadBoolean(markInfo, "Marked"),
            markInfo is null ? null : TryReadBoolean(markInfo, "Suspects"),
            markInfo is null ? null : TryReadBoolean(markInfo, "UserProperties"),
            structTreeRootObjectNumber,
            parentTreeObjectNumber,
            structTreeRoot is null ? null : TryReadInteger(structTreeRoot, "ParentTreeNextKey"),
            structTreeRoot is null ? EmptyReadOnlyDictionary() : ReadRoleMap(structTreeRoot),
            structTreeRoot is null ? Array.Empty<int>() : ReadStructureElementReferences(structTreeRoot.Items.TryGetValue("K", out PdfObject? kids) ? kids : null),
            parentTree is null ? Array.Empty<int>() : ReadParentTreeIndexes(parentTree),
            ReadStructureElements());
    }

    private bool? TryReadBoolean(PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            ResolveObject(value) is PdfBoolean boolean
            ? boolean.Value
            : null;
    }

    private System.Collections.ObjectModel.ReadOnlyDictionary<string, string> ReadRoleMap(PdfDictionary structTreeRoot) {
        PdfDictionary? roleMap = ResolveDict(structTreeRoot.Items.TryGetValue("RoleMap", out PdfObject? roleMapObject) ? roleMapObject : null);
        if (roleMap is null || roleMap.Items.Count == 0) {
            return EmptyReadOnlyDictionary();
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var entry in roleMap.Items) {
            if (TryFormatSimpleValue(entry.Value, out string? value) && !string.IsNullOrEmpty(value)) {
                values[entry.Key] = value!;
            }
        }

        return values.Count == 0 ? EmptyReadOnlyDictionary() : new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(values);
    }

    private IReadOnlyList<PdfStructureElementInfo> ReadStructureElements() {
        var elements = new List<PdfStructureElementInfo>();
        foreach (var item in _objects.OrderBy(entry => entry.Key)) {
            if (item.Value.Value is not PdfDictionary dictionary ||
                TryReadName(dictionary, "Type") != "StructElem") {
                continue;
            }

            int markedContentReferenceCount = 0;
            int objectReferenceCount = 0;
            IReadOnlyList<int> childElementObjectNumbers = dictionary.Items.TryGetValue("K", out PdfObject? kids)
                ? ReadStructureChildren(kids, ref markedContentReferenceCount, ref objectReferenceCount)
                : Array.Empty<int>();

            elements.Add(new PdfStructureElementInfo(
                item.Key,
                TryReadName(dictionary, "S"),
                ReadReferenceObjectNumber(dictionary, "P"),
                ReadReferenceObjectNumber(dictionary, "Pg"),
                TryReadText(dictionary, "Lang"),
                TryReadText(dictionary, "Alt"),
                childElementObjectNumbers,
                markedContentReferenceCount,
                objectReferenceCount));
        }

        return elements.Count == 0 ? Array.Empty<PdfStructureElementInfo>() : elements.AsReadOnly();
    }

    private IReadOnlyList<int> ReadStructureElementReferences(PdfObject? obj) {
        int markedContentReferenceCount = 0;
        int objectReferenceCount = 0;
        return ReadStructureChildren(obj, ref markedContentReferenceCount, ref objectReferenceCount, onlyStructureReferences: true);
    }

    private IReadOnlyList<int> ReadStructureChildren(PdfObject? obj, ref int markedContentReferenceCount, ref int objectReferenceCount, bool onlyStructureReferences = false) {
        var childObjectNumbers = new List<int>();
        AddStructureChildData(obj, childObjectNumbers, ref markedContentReferenceCount, ref objectReferenceCount, onlyStructureReferences);
        return childObjectNumbers.Count == 0 ? Array.Empty<int>() : childObjectNumbers.AsReadOnly();
    }

    private void AddStructureChildData(PdfObject? obj, List<int> childObjectNumbers, ref int markedContentReferenceCount, ref int objectReferenceCount, bool onlyStructureReferences) {
        PdfObject? resolved = ResolveObject(obj);
        if (obj is PdfReference reference && IsStructElementReference(reference)) {
            AddUnique(childObjectNumbers, reference.ObjectNumber);
            return;
        }

        if (resolved is PdfArray array) {
            for (int i = 0; i < array.Items.Count; i++) {
                AddStructureChildData(array.Items[i], childObjectNumbers, ref markedContentReferenceCount, ref objectReferenceCount, onlyStructureReferences);
            }

            return;
        }

        if (onlyStructureReferences) {
            return;
        }

        if (resolved is PdfNumber number && TryGetNonNegativeInteger(number, out _)) {
            markedContentReferenceCount++;
            return;
        }

        if (resolved is not PdfDictionary dictionary) {
            return;
        }

        string? type = TryReadName(dictionary, "Type");
        if (string.Equals(type, "MCR", StringComparison.Ordinal)) {
            markedContentReferenceCount++;
            return;
        }

        if (string.Equals(type, "OBJR", StringComparison.Ordinal)) {
            objectReferenceCount++;
            return;
        }

        if (dictionary.Items.TryGetValue("K", out PdfObject? nestedKids)) {
            AddStructureChildData(nestedKids, childObjectNumbers, ref markedContentReferenceCount, ref objectReferenceCount, onlyStructureReferences);
        }
    }

    private bool IsStructElementReference(PdfReference reference) {
        return PdfObjectLookup.TryGet(_objects, reference, out PdfIndirectObject? indirect) &&
            indirect.Value is PdfDictionary dictionary &&
            TryReadName(dictionary, "Type") == "StructElem";
    }

    private static int? ReadReferenceObjectNumber(PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) && value is PdfReference reference
            ? reference.ObjectNumber
            : null;
    }

    private IReadOnlyList<int> ReadParentTreeIndexes(PdfDictionary parentTree) {
        var indexes = new List<int>();
        AddParentTreeIndexes(parentTree, indexes, new HashSet<int>());
        return indexes.Count == 0 ? Array.Empty<int>() : indexes.AsReadOnly();
    }

    private void AddParentTreeIndexes(PdfObject? treeObject, List<int> indexes, HashSet<int> visitedReferences) {
        if (treeObject is PdfReference reference) {
            if (!visitedReferences.Add(reference.ObjectNumber)) {
                return;
            }

            treeObject = ResolveObject(reference);
        }

        if (treeObject is not PdfDictionary tree) {
            return;
        }

        if (tree.Items.TryGetValue("Nums", out PdfObject? numsObject) &&
            ResolveArray(numsObject) is PdfArray nums) {
            AddParentTreeNums(nums, indexes);
        }

        if (!tree.Items.TryGetValue("Kids", out PdfObject? kidsObject) ||
            ResolveArray(kidsObject) is not PdfArray kids) {
            return;
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            AddParentTreeIndexes(kids.Items[i], indexes, visitedReferences);
        }
    }

    private void AddParentTreeNums(PdfArray nums, List<int> indexes) {
        for (int i = 0; i + 1 < nums.Items.Count; i += 2) {
            if (ResolveObject(nums.Items[i]) is PdfNumber number &&
                TryGetNonNegativeInteger(number, out int index)) {
                AddUnique(indexes, index);
            }
        }
    }

    private static void AddUnique(List<int> values, int value) {
        if (!values.Contains(value)) {
            values.Add(value);
        }
    }

    private static System.Collections.ObjectModel.ReadOnlyDictionary<string, string> EmptyReadOnlyDictionary() {
        return new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(new Dictionary<string, string>(0, StringComparer.Ordinal));
    }
}
