namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    /// <summary>Tagged PDF structure metadata discovered from /MarkInfo and /StructTreeRoot.</summary>
    public PdfTaggedContentInfo? TaggedContent => ReadLogicalContent(_taggedContent);

    /// <summary>True when a readable tagged-PDF structure tree was discovered.</summary>
    public bool HasTaggedContent => TaggedContent is not null;

    private PdfTaggedContentInfo? ExtractTaggedContent(System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
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
            structTreeRoot is null ? EmptyReadOnlyDictionary() : ReadRoleMap(structTreeRoot, cancellationToken),
            structTreeRoot is null ? Array.Empty<int>() : ReadStructureElementReferences(structTreeRoot.Items.TryGetValue("K", out PdfObject? kids) ? kids : null, cancellationToken),
            parentTree is null ? Array.Empty<int>() : ReadParentTreeIndexes(parentTree, cancellationToken),
            ReadStructureElements(cancellationToken));
    }

    private bool? TryReadBoolean(PdfDictionary dictionary, string key) {
        return dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            ResolveObject(value) is PdfBoolean boolean
            ? boolean.Value
            : null;
    }

    private System.Collections.ObjectModel.ReadOnlyDictionary<string, string> ReadRoleMap(PdfDictionary structTreeRoot, System.Threading.CancellationToken cancellationToken) {
        PdfDictionary? roleMap = ResolveDict(structTreeRoot.Items.TryGetValue("RoleMap", out PdfObject? roleMapObject) ? roleMapObject : null);
        if (roleMap is null || roleMap.Items.Count == 0) {
            return EmptyReadOnlyDictionary();
        }

        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var entry in roleMap.Items) {
            cancellationToken.ThrowIfCancellationRequested();
            if (TryFormatSimpleValue(entry.Value, out string? value) && !string.IsNullOrEmpty(value)) {
                values[entry.Key] = value!;
            }
        }

        return values.Count == 0 ? EmptyReadOnlyDictionary() : new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(values);
    }

    private IReadOnlyList<PdfStructureElementInfo> ReadStructureElements(System.Threading.CancellationToken cancellationToken) {
        var elements = new List<PdfStructureElementInfo>();
        var orderedElements = new SortedDictionary<int, PdfIndirectObject>();
        foreach (var item in _objects) {
            cancellationToken.ThrowIfCancellationRequested();
            if (item.Value.Value is PdfDictionary dictionary && TryReadName(dictionary, "Type") == "StructElem") {
                orderedElements.Add(item.Key, item.Value);
            }
        }

        foreach (var item in orderedElements) {
            cancellationToken.ThrowIfCancellationRequested();
            var dictionary = (PdfDictionary)item.Value.Value;

            int? pageObjectNumber = ReadReferenceObjectNumber(dictionary, "Pg");
            int objectReferenceCount = 0;
            var markedContentReferences = new List<PdfMarkedContentReference>();
            IReadOnlyList<int> childElementObjectNumbers = dictionary.Items.TryGetValue("K", out PdfObject? kids)
                ? ReadStructureChildren(kids, pageObjectNumber, markedContentReferences, ref objectReferenceCount, cancellationToken)
                : Array.Empty<int>();

            elements.Add(new PdfStructureElementInfo(
                item.Key,
                TryReadName(dictionary, "S"),
                ReadReferenceObjectNumber(dictionary, "P"),
                pageObjectNumber,
                TryReadText(dictionary, "Lang"),
                TryReadText(dictionary, "Alt"),
                childElementObjectNumbers,
                markedContentReferences.Count == 0 ? Array.Empty<PdfMarkedContentReference>() : markedContentReferences.AsReadOnly(),
                objectReferenceCount));
        }

        return elements.Count == 0 ? Array.Empty<PdfStructureElementInfo>() : elements.AsReadOnly();
    }

    private IReadOnlyList<int> ReadStructureElementReferences(PdfObject? obj, System.Threading.CancellationToken cancellationToken) {
        var markedContentReferences = new List<PdfMarkedContentReference>();
        int objectReferenceCount = 0;
        return ReadStructureChildren(obj, null, markedContentReferences, ref objectReferenceCount, cancellationToken, onlyStructureReferences: true);
    }

    private IReadOnlyList<int> ReadStructureChildren(
        PdfObject? obj,
        int? inheritedPageObjectNumber,
        List<PdfMarkedContentReference> markedContentReferences,
        ref int objectReferenceCount,
        System.Threading.CancellationToken cancellationToken,
        bool onlyStructureReferences = false) {
        var childObjectNumbers = new List<int>();
        AddStructureChildData(obj, inheritedPageObjectNumber, childObjectNumbers, markedContentReferences, ref objectReferenceCount, onlyStructureReferences, cancellationToken);
        return childObjectNumbers.Count == 0 ? Array.Empty<int>() : childObjectNumbers.AsReadOnly();
    }

    private void AddStructureChildData(
        PdfObject? obj,
        int? inheritedPageObjectNumber,
        List<int> childObjectNumbers,
        List<PdfMarkedContentReference> markedContentReferences,
        ref int objectReferenceCount,
        bool onlyStructureReferences,
        System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfObject? resolved = ResolveObject(obj);
        if (obj is PdfReference reference && IsStructElementReference(reference)) {
            AddUnique(childObjectNumbers, reference.ObjectNumber, cancellationToken);
            return;
        }

        if (resolved is PdfArray array) {
            for (int i = 0; i < array.Items.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                AddStructureChildData(array.Items[i], inheritedPageObjectNumber, childObjectNumbers, markedContentReferences, ref objectReferenceCount, onlyStructureReferences, cancellationToken);
            }

            return;
        }

        if (onlyStructureReferences) {
            return;
        }

        if (resolved is PdfNumber number && TryGetNonNegativeInteger(number, out int markedContentId)) {
            markedContentReferences.Add(new PdfMarkedContentReference(inheritedPageObjectNumber, markedContentId));
            return;
        }

        if (resolved is not PdfDictionary dictionary) {
            return;
        }

        string? type = TryReadName(dictionary, "Type");
        bool isMarkedContentReference = string.Equals(type, "MCR", StringComparison.Ordinal) ||
            (type is null && dictionary.Items.ContainsKey("MCID"));
        if (isMarkedContentReference) {
            int? pageObjectNumber = ReadReferenceObjectNumber(dictionary, "Pg") ?? inheritedPageObjectNumber;
            int? contentStreamObjectNumber = ReadReferenceObjectNumber(dictionary, "Stm");
            if (dictionary.Items.TryGetValue("MCID", out PdfObject? mcidObject) &&
                ResolveObject(mcidObject) is PdfNumber mcidNumber &&
                TryGetNonNegativeInteger(mcidNumber, out int referencedMcid)) {
                markedContentReferences.Add(new PdfMarkedContentReference(pageObjectNumber, referencedMcid, contentStreamObjectNumber));
            }
            return;
        }

        if (string.Equals(type, "OBJR", StringComparison.Ordinal)) {
            objectReferenceCount++;
            return;
        }

        if (dictionary.Items.TryGetValue("K", out PdfObject? nestedKids)) {
            int? nestedPageObjectNumber = ReadReferenceObjectNumber(dictionary, "Pg") ?? inheritedPageObjectNumber;
            AddStructureChildData(nestedKids, nestedPageObjectNumber, childObjectNumbers, markedContentReferences, ref objectReferenceCount, onlyStructureReferences, cancellationToken);
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

    private IReadOnlyList<int> ReadParentTreeIndexes(PdfDictionary parentTree, System.Threading.CancellationToken cancellationToken) {
        var indexes = new List<int>();
        AddParentTreeIndexes(parentTree, indexes, new HashSet<int>(), cancellationToken);
        return indexes.Count == 0 ? Array.Empty<int>() : indexes.AsReadOnly();
    }

    private void AddParentTreeIndexes(PdfObject? treeObject, List<int> indexes, HashSet<int> visitedReferences, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
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
            AddParentTreeNums(nums, indexes, cancellationToken);
        }

        if (!tree.Items.TryGetValue("Kids", out PdfObject? kidsObject) ||
            ResolveArray(kidsObject) is not PdfArray kids) {
            return;
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            AddParentTreeIndexes(kids.Items[i], indexes, visitedReferences, cancellationToken);
        }
    }

    private void AddParentTreeNums(PdfArray nums, List<int> indexes, System.Threading.CancellationToken cancellationToken) {
        for (int i = 0; i + 1 < nums.Items.Count; i += 2) {
            cancellationToken.ThrowIfCancellationRequested();
            if (ResolveObject(nums.Items[i]) is PdfNumber number &&
                TryGetNonNegativeInteger(number, out int index)) {
                AddUnique(indexes, index, cancellationToken);
            }
        }
    }

    private static void AddUnique(List<int> values, int value, System.Threading.CancellationToken cancellationToken) {
        foreach (int existing in values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (existing == value) return;
        }
        values.Add(value);
    }

    private static System.Collections.ObjectModel.ReadOnlyDictionary<string, string> EmptyReadOnlyDictionary() {
        return new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(new Dictionary<string, string>(0, StringComparer.Ordinal));
    }
}
