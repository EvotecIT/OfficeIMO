using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfTaggedPageOrganizer {
    private static readonly HashSet<string> ImportableStructureRootKeys = new(new[] {
        "Type", "K", "ParentTree", "ParentTreeNextKey", "IDTree", "RoleMap", "ClassMap", "Namespaces", "AF"
    }, StringComparer.Ordinal);
    internal static void MergeImportedStructure(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog,
        PdfDictionary importedCatalog, IEnumerable<int> importedNumbers, int[] originalPages, int insertBeforePage, CancellationToken token) {
        if (!importedCatalog.Items.TryGetValue("StructTreeRoot", out var sourceRootValue) || PdfObjectLookup.Resolve(objects, sourceRootValue) is not PdfDictionary sourceRoot) return;
        // Catalog language inheritance ends when roots move to another catalog.
        if (importedCatalog.Items.TryGetValue("Lang", out var sourceLanguage) && sourceRoot.Items.TryGetValue("K", out var sourceChildrenValue)) {
            var sourceRoots = PdfObjectLookup.Resolve(objects, sourceChildrenValue) is PdfArray sourceArray ? sourceArray.Items : new List<PdfObject> { sourceChildrenValue };
            foreach (var child in sourceRoots) if (PdfObjectLookup.Resolve(objects, child) is PdfDictionary element && !element.Items.ContainsKey("Lang"))
                element.Items["Lang"] = sourceLanguage;
        }
        if (!catalog.Items.TryGetValue("StructTreeRoot", out var targetRootValue) || PdfObjectLookup.Resolve(objects, targetRootValue) is not PdfDictionary targetRoot) {
            catalog.Items["StructTreeRoot"] = sourceRootValue;
            if (importedCatalog.Items.TryGetValue("MarkInfo", out var mark)) catalog.Items["MarkInfo"] = mark;
            if (!catalog.Items.ContainsKey("Lang") && importedCatalog.Items.TryGetValue("Lang", out var language)) catalog.Items["Lang"] = language;
            return;
        }
        if (sourceRoot.Items.Keys.Any(key => !ImportableStructureRootKeys.Contains(key)))
            throw new NotSupportedException("The incoming structure root contains an extension that requires an explicit merge policy.");
        var targetEntries = targetRoot.Items.TryGetValue("ParentTree", out var targetParents) ? ReadNumberTree(objects, targetParents, token) : new SortedDictionary<int, PdfObject>();
        var sourceEntries = sourceRoot.Items.TryGetValue("ParentTree", out var sourceParents) ? ReadNumberTree(objects, sourceParents, token) : new SortedDictionary<int, PdfObject>();
        int nextKey = targetEntries.Count == 0 ? 0 : checked(targetEntries.Keys.Max() + 1);
        var keys = new Dictionary<int, int>();
        foreach (var pair in sourceEntries) { int key = nextKey; nextKey = checked(nextKey + 1); keys.Add(pair.Key, key); targetEntries.Add(key, pair.Value); }
        int[] imported = importedNumbers.ToArray();
        foreach (int number in imported) {
            token.ThrowIfCancellationRequested();
            var dictionary = objects[number].Value is PdfStream stream ? stream.Dictionary : objects[number].Value as PdfDictionary;
            if (dictionary == null) continue;
            foreach (string name in new[] { "StructParents", "StructParent" }) if (dictionary.Get<PdfNumber>(name) is PdfNumber value) {
                if (!keys.TryGetValue(CheckedKey(value), out int mapped)) throw new InvalidOperationException("An imported tagged object has no parent-tree entry.");
                dictionary.Items[name] = new PdfNumber(mapped);
            }
        }
        var nums = new PdfArray(); foreach (var pair in targetEntries) { nums.Items.Add(new PdfNumber(pair.Key)); nums.Items.Add(pair.Value); }
        var parentTree = new PdfDictionary(); parentTree.Items["Nums"] = nums; targetRoot.Items["ParentTree"] = parentTree;
        targetRoot.Items["ParentTreeNextKey"] = new PdfNumber(nextKey);
        string prefix = UniqueStructureIdPrefix(objects, "import" + (-imported.Min()).ToString(System.Globalization.CultureInfo.InvariantCulture) + ":");
        var roles = MergeMap("RoleMap"); var classes = MergeMap("ClassMap");
        Dictionary<string, string> MergeMap(string name) {
            var renames = new Dictionary<string, string>(StringComparer.Ordinal);
            if (!sourceRoot.Items.TryGetValue(name, out var value) || PdfObjectLookup.Resolve(objects, value) is not PdfDictionary sourceMap) return renames;
            var combined = new PdfDictionary();
            if (targetRoot.Items.TryGetValue(name, out var old) && PdfObjectLookup.Resolve(objects, old) is PdfDictionary targetMap)
                foreach (var pair in targetMap.Items) combined.Items[pair.Key] = pair.Value;
            foreach (var pair in sourceMap.Items) {
                string renamed = prefix + pair.Key; while (combined.Items.ContainsKey(renamed)) renamed = "_" + renamed;
                renames.Add(pair.Key, renamed); combined.Items[renamed] = pair.Value;
            }
            if (name == "RoleMap") foreach (string key in renames.Values)
                if (combined.Items[key] is PdfName role && renames.TryGetValue(role.Name, out string? mapped)) combined.Items[key] = new PdfName(mapped);
            targetRoot.Items[name] = combined; return renames;
        }
        var visited = new HashSet<int>(); var ids = new Dictionary<string, PdfStringObj>(StringComparer.Ordinal);
        var activeArrays = new HashSet<PdfArray>();
        var renamedClasses = new Dictionary<PdfArray, PdfReference>();
        int lastClassObjectNumber = objects.Keys.Max();
        void Rename(PdfObject value, int depth) {
            token.ThrowIfCancellationRequested();
            RequireStructureDepth(depth);
            var resolved = PdfObjectLookup.Resolve(objects, value);
            if (resolved is PdfArray array) {
                if (!activeArrays.Add(array)) throw new InvalidOperationException("The imported structure tree contains a cyclic child array.");
                foreach (var item in array.Items) Rename(item, depth + 1);
                activeArrays.Remove(array);
                return;
            }
            if (value is not PdfReference reference || !visited.Add(reference.ObjectNumber) || resolved is not PdfDictionary dictionary) return;
            if (!dictionary.Items.ContainsKey("NS") && dictionary.Get<PdfName>("S") is PdfName role && roles.TryGetValue(role.Name, out string? renamed)) dictionary.Items["S"] = new PdfName(renamed);
            if (dictionary.Items.TryGetValue("C", out var c)) dictionary.Items["C"] = RenameClasses(c);
            if (dictionary.Items.TryGetValue("ID", out var id) && PdfObjectLookup.Resolve(objects, id) is PdfStringObj text) {
                var mapped = PrefixStructureId(prefix, text); dictionary.Items["ID"] = mapped;
                ids.Add(Convert.ToBase64String(text.RawBytes), mapped);
            }
            if (dictionary.Items.TryGetValue("K", out var children)) Rename(children, depth + 1);
        }
        PdfObject RenameClasses(PdfObject value) {
            token.ThrowIfCancellationRequested();
            var resolved = PdfObjectLookup.Resolve(objects, value);
            if (resolved is PdfName name && classes.TryGetValue(name.Name, out string? renamed)) return new PdfName(renamed);
            if (resolved is PdfArray array) {
                if (renamedClasses.TryGetValue(array, out var previous)) return previous;
                var result = new PdfArray();
                foreach (var item in array.Items) {
                    token.ThrowIfCancellationRequested();
                    if (PdfObjectLookup.Resolve(objects, item) is PdfArray)
                        throw new InvalidOperationException("Nested structure class arrays are not supported.");
                    result.Items.Add(RenameClasses(item));
                }
                if (value is not PdfReference) return result;
                // Preserve sharing without changing an array that another catalog value may also use.
                int number = lastClassObjectNumber = checked(lastClassObjectNumber + 1);
                var reference = new PdfReference(number, 0);
                objects.Add(number, new PdfIndirectObject(number, 0, result));
                renamedClasses.Add(array, reference);
                return reference;
            }
            return value;
        }
        if (sourceRoot.Items.TryGetValue("K", out var sourceKids)) Rename(sourceKids, 0);
        RemapStructureDestinations(objects, imported, ids, token);
        foreach (string name in new[] { "Namespaces", "AF" }) if (sourceRoot.Items.TryGetValue(name, out var value)) {
            var combined = new PdfArray();
            if (targetRoot.Items.TryGetValue(name, out var old) && PdfObjectLookup.Resolve(objects, old) is PdfArray existing) combined.Items.AddRange(existing.Items);
            if (PdfObjectLookup.Resolve(objects, value) is not PdfArray added) throw new InvalidOperationException("The imported structure root has an invalid " + name + " array.");
            combined.Items.AddRange(added.Items); targetRoot.Items[name] = combined;
        }
        PdfDictionary parent = targetRoot; PdfObject parentValue = targetRootValue;
        var targetChildren = Children(targetRoot); var sourceChildren = Children(sourceRoot);
        if (targetChildren.Count == 1 && PdfObjectLookup.Resolve(objects, targetChildren[0]) is PdfDictionary document && document.Get<PdfName>("S")?.Name == "Document") {
            parent = document; parentValue = targetChildren[0]; targetChildren = Children(document);
            foreach (var child in sourceChildren) if (PdfObjectLookup.Resolve(objects, child) is PdfDictionary item && item.Get<PdfName>("S")?.Name == "Document") {
                item.Items["S"] = new PdfName("Part");
                if (!item.Items.ContainsKey("Lang") && importedCatalog.Items.TryGetValue("Lang", out var language)) item.Items["Lang"] = language;
            }
        }
        foreach (var child in sourceChildren) {
            if (PdfObjectLookup.Resolve(objects, child) is not PdfDictionary item || child is not PdfReference)
                throw new InvalidOperationException("Imported root structure children must be indirect structure elements.");
            item.Items["P"] = parentValue;
        }
        var pageIndexes = originalPages.Select((page, index) => (page, index)).ToDictionary(item => item.page, item => item.index + 1);
        int insertion = targetChildren.FindIndex(child => FirstPage(child, parent.Get<PdfReference>("Pg")?.ObjectNumber, new HashSet<int>(), 0) >= insertBeforePage);
        if (insertion < 0) insertion = targetChildren.Count;
        targetChildren.InsertRange(insertion, sourceChildren); var combinedKids = new PdfArray(); combinedKids.Items.AddRange(targetChildren); parent.Items["K"] = combinedKids;
        RebuildStructureIds(objects, targetRoot, token);

        List<PdfObject> Children(PdfDictionary dictionary) {
            if (!dictionary.Items.TryGetValue("K", out var value)) return new List<PdfObject>();
            return PdfObjectLookup.Resolve(objects, value) is PdfArray array ? array.Items.ToList() : new List<PdfObject> { value };
        }
        int FirstPage(PdfObject value, int? inherited, HashSet<int> seen, int depth) {
            token.ThrowIfCancellationRequested();
            RequireStructureDepth(depth);
            var resolved = PdfObjectLookup.Resolve(objects, value);
            if (value is PdfReference reference && !seen.Add(reference.ObjectNumber)) throw new InvalidOperationException("The target structure tree is cyclic or shared.");
            if (resolved is PdfArray array) return array.Items.Count == 0 ? int.MaxValue : array.Items.Min(item => FirstPage(item, inherited, seen, depth + 1));
            if (resolved is PdfNumber) return inherited.HasValue && pageIndexes.TryGetValue(inherited.Value, out int index) ? index : int.MaxValue;
            if (resolved is not PdfDictionary dictionary) return int.MaxValue;
            int? page = dictionary.Get<PdfReference>("Pg")?.ObjectNumber ?? inherited;
            int own = page.HasValue && pageIndexes.TryGetValue(page.Value, out int ownIndex) ? ownIndex : int.MaxValue;
            return dictionary.Items.TryGetValue("K", out var kids) ? Math.Min(own, FirstPage(kids, page, seen, depth + 1)) : own;
        }
    }
}
