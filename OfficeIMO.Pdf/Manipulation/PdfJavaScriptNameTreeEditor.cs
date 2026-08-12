namespace OfficeIMO.Pdf;

/// <summary>Applies exact-key commands while preserving untouched JavaScript name-tree values.</summary>
internal static partial class PdfJavaScriptNameTreeEditor {
    internal static void Rewrite(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        IReadOnlyList<PdfJavaScriptEditSession.EditCommand> commands,
        PdfReadLimits limits) {
        if (!security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? root) ||
            root.Value is not PdfDictionary catalog) {
            throw new InvalidOperationException("PDF catalog is not readable.");
        }

        bool hasNames = catalog.Items.TryGetValue("Names", out PdfObject? namesObject);
        PdfDictionary? names = ResolveDictionary(objects, hasNames ? namesObject : null);
        if (hasNames && names is null) {
            throw new InvalidDataException("The PDF catalog /Names entry is not a readable name-tree dictionary.");
        }

        int lastClear = -1;
        for (int i = 0; i < commands.Count; i++) {
            if (commands[i].Kind == PdfJavaScriptEditSession.EditKind.Clear) lastClear = i;
        }

        var entries = new List<NameTreeEntry>();
        if (lastClear < 0 && names is not null && names.Items.TryGetValue("JavaScript", out PdfObject? treeObject)) {
            int traversedNodes = 0;
            CollectEntries(
                objects,
                treeObject,
                entries,
                new HashSet<(int ObjectNumber, int Generation)>(),
                0,
                ref traversedNodes,
                limits);
        }

        int firstCommand = lastClear < 0 ? 0 : lastClear + 1;
        for (int i = firstCommand; i < commands.Count; i++) {
            ApplyCommand(objects, entries, commands[i]);
        }

        if (entries.Count > limits.MaxJavaScripts) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.JavaScripts, limits.MaxJavaScripts, entries.Count);
        }
        if (entries.Count == 0) {
            if (names is not null) {
                names.Items.Remove("JavaScript");
                if (names.Items.Count == 0) catalog.Items.Remove("Names");
            }
            return;
        }

        if (names is null) {
            names = new PdfDictionary();
            catalog.Items["Names"] = names;
        }

        entries.Sort(NameTreeEntryComparer.Instance);
        var values = new PdfArray();
        for (int i = 0; i < entries.Count; i++) {
            values.Items.Add(new PdfStringObj(entries[i].KeyBytes, useTextStringEncoding: true));
            values.Items.Add(entries[i].Value);
        }
        var tree = new PdfDictionary();
        tree.Items["Names"] = values;
        names.Items["JavaScript"] = tree;
    }

    private static void ApplyCommand(
        Dictionary<int, PdfIndirectObject> objects,
        List<NameTreeEntry> entries,
        PdfJavaScriptEditSession.EditCommand command) {
        if (command.Kind == PdfJavaScriptEditSession.EditKind.Clear) {
            entries.Clear();
            return;
        }

        var matches = new List<int>();
        for (int i = 0; i < entries.Count; i++) {
            if (string.Equals(entries[i].Name, command.Name, StringComparison.Ordinal)) matches.Add(i);
        }
        if (command.Kind == PdfJavaScriptEditSession.EditKind.Remove) {
            for (int i = matches.Count - 1; i >= 0; i--) entries.RemoveAt(matches[i]);
            return;
        }

        if (matches.Count > 1) {
            throw new InvalidDataException("The PDF JavaScript name tree contains duplicate entries for the edited key.");
        }

        PdfDictionary replacement;
        if (matches.Count == 1) {
            PdfObject? existing = PdfObjectLookup.Resolve(objects, entries[matches[0]].Value);
            if (existing is not PdfDictionary action ||
                !action.Items.TryGetValue("S", out PdfObject? actionTypeObject) ||
                PdfObjectLookup.Resolve(objects, actionTypeObject) is not PdfName actionType ||
                actionType.Name != "JavaScript") {
                throw new InvalidDataException("The edited PDF JavaScript name does not reference a readable JavaScript action dictionary.");
            }
            replacement = CloneDictionary(action);
            replacement.Items["JS"] = new PdfStringObj(command.EncodedScript!, useTextStringEncoding: true);
            NameTreeEntry existingEntry = entries[matches[0]];
            entries[matches[0]] = new NameTreeEntry(
                existingEntry.KeyBytes,
                command.Name!,
                replacement,
                existingEntry.OriginalPosition);
            return;
        }

        replacement = new PdfDictionary();
        replacement.Items["S"] = new PdfName("JavaScript");
        replacement.Items["JS"] = new PdfStringObj(command.EncodedScript!, useTextStringEncoding: true);
        entries.Add(new NameTreeEntry(command.EncodedName!, command.Name!, replacement, entries.Count));
    }

    private static void CollectEntries(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject treeObject,
        List<NameTreeEntry> entries,
        HashSet<(int ObjectNumber, int Generation)> visited,
        int depth,
        ref int traversedNodes,
        PdfReadLimits limits) {
        if (depth > limits.MaxNameTreeDepth) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.NameTreeDepth, limits.MaxNameTreeDepth, depth);
        }
        if (treeObject is PdfReference reference) {
            if (!visited.Add((reference.ObjectNumber, reference.Generation))) {
                throw new InvalidDataException("The PDF JavaScript name tree contains a reference cycle.");
            }
            traversedNodes++;
            if (traversedNodes > limits.MaxNameTreeNodes) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.NameTreeNodes, limits.MaxNameTreeNodes, traversedNodes);
            }
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) {
                throw new InvalidDataException("The PDF JavaScript name tree contains an unresolved reference.");
            }
            treeObject = indirect.Value;
        }
        if (treeObject is not PdfDictionary tree) {
            throw new InvalidDataException("The PDF catalog JavaScript entry is not a readable name tree.");
        }
        foreach (string key in tree.Items.Keys) {
            if (key != "Names" && key != "Kids" && key != "Limits") {
                throw new InvalidDataException("The PDF JavaScript name tree contains unsupported extension data that cannot be normalized losslessly.");
            }
        }
        bool hasValues = tree.Items.TryGetValue("Names", out PdfObject? valuesObject);
        bool hasKids = tree.Items.TryGetValue("Kids", out PdfObject? kidsObject);
        if (hasValues && hasKids) {
            throw new InvalidDataException("The PDF JavaScript name tree node contains both /Names and /Kids.");
        }
        if (hasValues) {
            if (PdfObjectLookup.Resolve(objects, valuesObject) is not PdfArray values || (values.Items.Count & 1) != 0) {
                throw new InvalidDataException("The PDF JavaScript name tree contains an invalid /Names array.");
            }
            for (int i = 0; i < values.Items.Count; i += 2) {
                if (PdfObjectLookup.Resolve(objects, values.Items[i]) is not PdfStringObj key) {
                    throw new InvalidDataException("The PDF JavaScript name tree contains a non-string key.");
                }
                string? name = PdfJavaScriptStringEncoding.TryDecode(key.RawBytes, out string decoded) ? decoded : null;
                entries.Add(new NameTreeEntry(key.RawBytes, name, values.Items[i + 1], entries.Count));
                if (entries.Count > limits.MaxJavaScripts) {
                    throw PdfReadLimitException.Create(PdfReadLimitKind.JavaScripts, limits.MaxJavaScripts, entries.Count);
                }
            }
        }
        if (hasKids) {
            if (PdfObjectLookup.Resolve(objects, kidsObject) is not PdfArray kids) {
                throw new InvalidDataException("The PDF JavaScript name tree contains an invalid /Kids array.");
            }
            for (int i = 0; i < kids.Items.Count; i++) {
                CollectEntries(objects, kids.Items[i], entries, visited, depth + 1, ref traversedNodes, limits);
            }
        }
    }

    private static PdfDictionary CloneDictionary(PdfDictionary source) {
        var clone = new PdfDictionary();
        foreach (KeyValuePair<string, PdfObject> item in source.Items) clone.Items[item.Key] = item.Value;
        return clone;
    }

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) =>
        PdfObjectLookup.Resolve(objects, value) as PdfDictionary;

    private sealed class NameTreeEntry {
        internal NameTreeEntry(byte[] keyBytes, string? name, PdfObject value, int originalPosition) {
            KeyBytes = (byte[])keyBytes.Clone(); Name = name; Value = value; OriginalPosition = originalPosition;
        }
        internal byte[] KeyBytes { get; }
        internal string? Name { get; }
        internal PdfObject Value { get; }
        internal int OriginalPosition { get; }
    }

    private sealed class NameTreeEntryComparer : IComparer<NameTreeEntry> {
        internal static NameTreeEntryComparer Instance { get; } = new NameTreeEntryComparer();
        public int Compare(NameTreeEntry? x, NameTreeEntry? y) {
            if (ReferenceEquals(x, y)) return 0;
            if (x is null) return -1;
            if (y is null) return 1;
            int count = Math.Min(x.KeyBytes.Length, y.KeyBytes.Length);
            for (int i = 0; i < count; i++) {
                int comparison = x.KeyBytes[i].CompareTo(y.KeyBytes[i]);
                if (comparison != 0) return comparison;
            }
            int lengthComparison = x.KeyBytes.Length.CompareTo(y.KeyBytes.Length);
            return lengthComparison != 0 ? lengthComparison : x.OriginalPosition.CompareTo(y.OriginalPosition);
        }
    }
}
