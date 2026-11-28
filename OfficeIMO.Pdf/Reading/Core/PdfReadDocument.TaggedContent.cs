namespace OfficeIMO.Pdf;

internal sealed partial class PdfReadDocument {
    private const int MaxReadTaggedContentDepth = 256;

    /// <summary>Tagged PDF structure metadata discovered from /MarkInfo and /StructTreeRoot.</summary>
    public PdfTaggedContentInfo? TaggedContent => ReadLogicalFeature(_taggedContentFeature);

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
            if (TryFormatSimpleValue(entry.Value, out string? value, cancellationToken) && !string.IsNullOrEmpty(value)) {
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

        PdfDictionary? classMap = null;
        if (orderedElements.Count > 0 &&
            FindCatalog() is PdfDictionary catalog &&
            ResolveDict(catalog.Items.TryGetValue("StructTreeRoot", out PdfObject? rootObject) ? rootObject : null) is PdfDictionary root) {
            classMap = ResolveDict(root.Items.TryGetValue("ClassMap", out PdfObject? classMapObject) ? classMapObject : null);
        }

        foreach (var item in orderedElements) {
            cancellationToken.ThrowIfCancellationRequested();
            var dictionary = (PdfDictionary)item.Value.Value;
            string? structureType = TryReadName(dictionary, "S");
            int rowSpan = 1;
            int columnSpan = 1;
            if (structureType is not null &&
                (dictionary.Items.ContainsKey("A") || dictionary.Items.ContainsKey("C"))) {
                ReadTableCellSpans(dictionary, classMap, ref rowSpan, ref columnSpan, cancellationToken);
            }

            int? pageObjectNumber = ReadReferenceObjectNumber(dictionary, "Pg");
            int objectReferenceCount = 0;
            var markedContentReferences = new List<PdfMarkedContentReference>();
            IReadOnlyList<int> childElementObjectNumbers = dictionary.Items.TryGetValue("K", out PdfObject? kids)
                ? ReadStructureChildren(kids, pageObjectNumber, markedContentReferences, ref objectReferenceCount, cancellationToken)
                : Array.Empty<int>();

            elements.Add(new PdfStructureElementInfo(
                item.Key,
                structureType,
                ReadReferenceObjectNumber(dictionary, "P"),
                pageObjectNumber,
                TryReadText(dictionary, "Lang"),
                TryReadText(dictionary, "Alt"),
                childElementObjectNumbers,
                markedContentReferences.Count == 0 ? Array.Empty<PdfMarkedContentReference>() : markedContentReferences.AsReadOnly(),
                objectReferenceCount,
                rowSpan,
                columnSpan));
        }

        return elements.Count == 0 ? Array.Empty<PdfStructureElementInfo>() : elements.AsReadOnly();
    }

    /// <summary>
    /// Reads RowSpan and ColSpan from Table attribute objects. Classes from /C apply first;
    /// attributes on the element itself (/A) take precedence, as ISO 32000 requires.
    /// </summary>
    private void ReadTableCellSpans(
        PdfDictionary element,
        PdfDictionary? classMap,
        ref int rowSpan,
        ref int columnSpan,
        System.Threading.CancellationToken cancellationToken) {
        if (classMap is not null && element.Items.TryGetValue("C", out PdfObject? classesObject)) {
            PdfObject? classes = ResolveObject(classesObject);
            IReadOnlyList<PdfObject> names = classes is PdfArray classArray ? classArray.Items : classes is PdfName ? new[] { classes } : Array.Empty<PdfObject>();
            for (int index = 0; index < names.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (ResolveObject(names[index]) is PdfName name &&
                    classMap.Items.TryGetValue(name.Name, out PdfObject? classAttributes)) {
                    ApplyTableAttributeSpans(classAttributes, ref rowSpan, ref columnSpan, cancellationToken);
                }
            }
        }

        if (element.Items.TryGetValue("A", out PdfObject? attributes)) {
            ApplyTableAttributeSpans(attributes, ref rowSpan, ref columnSpan, cancellationToken);
        }
    }

    private void ApplyTableAttributeSpans(
        PdfObject? attributes,
        ref int rowSpan,
        ref int columnSpan,
        System.Threading.CancellationToken cancellationToken) {
        PdfObject? resolved = ResolveObject(attributes);
        if (resolved is PdfArray array) {
            // Revision numbers may follow attribute objects; only dictionaries carry values.
            for (int index = 0; index < array.Items.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (AsAttributeDictionary(ResolveObject(array.Items[index])) is PdfDictionary item) ApplyTableAttributeDictionary(item, ref rowSpan, ref columnSpan);
            }
            return;
        }
        if (AsAttributeDictionary(resolved) is PdfDictionary dictionary) ApplyTableAttributeDictionary(dictionary, ref rowSpan, ref columnSpan);
    }

    /// <summary>An attribute object is a dictionary or a stream whose dictionary holds the attributes.</summary>
    private static PdfDictionary? AsAttributeDictionary(PdfObject? value) =>
        value as PdfDictionary ?? (value as PdfStream)?.Dictionary;

    private void ApplyTableAttributeDictionary(PdfDictionary attributes, ref int rowSpan, ref int columnSpan) {
        if (!string.Equals(TryReadName(attributes, "O"), "Table", StringComparison.Ordinal)) return;
        if (attributes.Items.TryGetValue("RowSpan", out PdfObject? rowObject) &&
            ResolveObject(rowObject) is PdfNumber rowNumber &&
            TryGetNonNegativeInteger(rowNumber, out int rows) && rows > 0) {
            rowSpan = rows;
        }
        if (attributes.Items.TryGetValue("ColSpan", out PdfObject? columnObject) &&
            ResolveObject(columnObject) is PdfNumber columnNumber &&
            TryGetNonNegativeInteger(columnNumber, out int columns) && columns > 0) {
            columnSpan = columns;
        }
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
        AddStructureChildData(obj, inheritedPageObjectNumber, childObjectNumbers, markedContentReferences, ref objectReferenceCount, onlyStructureReferences, new HashSet<PdfObject>(), 0, cancellationToken);
        return childObjectNumbers.Count == 0 ? Array.Empty<int>() : childObjectNumbers.AsReadOnly();
    }

    private void AddStructureChildData(
        PdfObject? obj,
        int? inheritedPageObjectNumber,
        List<int> childObjectNumbers,
        List<PdfMarkedContentReference> markedContentReferences,
        ref int objectReferenceCount,
        bool onlyStructureReferences,
        HashSet<PdfObject> ancestorContainers,
        int depth,
        System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        RequireTaggedContentDepth(depth);
        PdfObject? resolved = ResolveObject(obj);
        if (obj is PdfReference reference && IsStructElementReference(reference)) {
            AddUnique(childObjectNumbers, reference.ObjectNumber, cancellationToken);
            return;
        }

        if (resolved is PdfArray array) {
            // A repeated sibling may legitimately reuse an array; only ancestors form a cycle.
            if (!ancestorContainers.Add(array)) return;
            try {
                for (int i = 0; i < array.Items.Count; i++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    AddStructureChildData(array.Items[i], inheritedPageObjectNumber, childObjectNumbers, markedContentReferences, ref objectReferenceCount, onlyStructureReferences, ancestorContainers, depth + 1, cancellationToken);
                }
            } finally {
                ancestorContainers.Remove(array);
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
            if (!ancestorContainers.Add(dictionary)) return;
            try {
                int? nestedPageObjectNumber = ReadReferenceObjectNumber(dictionary, "Pg") ?? inheritedPageObjectNumber;
                AddStructureChildData(nestedKids, nestedPageObjectNumber, childObjectNumbers, markedContentReferences, ref objectReferenceCount, onlyStructureReferences, ancestorContainers, depth + 1, cancellationToken);
            } finally {
                ancestorContainers.Remove(dictionary);
            }
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
        AddParentTreeIndexes(parentTree, indexes, new HashSet<int>(), 0, cancellationToken);
        return indexes.Count == 0 ? Array.Empty<int>() : indexes.AsReadOnly();
    }

    private void AddParentTreeIndexes(PdfObject? treeObject, List<int> indexes, HashSet<int> visitedReferences, int depth, System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        RequireTaggedContentDepth(depth);
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
            AddParentTreeIndexes(kids.Items[i], indexes, visitedReferences, depth + 1, cancellationToken);
        }
    }

    /// <summary>Bounds semantic tag traversal across indirect objects as well as directly nested syntax.</summary>
    private static void RequireTaggedContentDepth(int depth) {
        if (depth > MaxReadTaggedContentDepth) {
            throw new InvalidOperationException("Tagged PDF reading exceeded the structure nesting limit.");
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
