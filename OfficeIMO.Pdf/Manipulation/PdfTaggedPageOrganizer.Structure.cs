using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfTaggedPageOrganizer {
    /// <summary>Bounds semantic traversal across indirect objects as well as directly nested syntax.</summary>
    private static void RequireStructureDepth(int depth) {
        if (depth > 256) throw new InvalidOperationException("Tagged page organization exceeded the structure nesting limit.");
    }

    private static string UniqueStructureIdPrefix(Dictionary<int, PdfIndirectObject> objects, string preferred) {
        var existing = objects.Values.Select(item => item.Value).OfType<PdfDictionary>()
            .Where(dictionary => dictionary.Items.ContainsKey("ID"))
            .Select(dictionary => PdfObjectLookup.Resolve(objects, dictionary.Items["ID"]))
            .OfType<PdfStringObj>().Select(id => id.RawBytes).ToArray();
        string prefix = preferred;
        while (existing.Any(id => id.Take(prefix.Length).SequenceEqual(System.Text.Encoding.ASCII.GetBytes(prefix)))) prefix = "_" + prefix;
        return prefix;
    }

    private static PdfStringObj PrefixStructureId(string prefix, PdfStringObj id) =>
        new(System.Text.Encoding.ASCII.GetBytes(prefix).Concat(id.RawBytes).ToArray());

    private static void RemapStructureDestinations(Dictionary<int, PdfIndirectObject> objects, IEnumerable<int> numbers,
        Dictionary<string, PdfStringObj> ids, CancellationToken token) {
        var visited = new HashSet<PdfObject>();
        void Visit(PdfObject value, int depth) {
            token.ThrowIfCancellationRequested();
            RequireStructureDepth(depth);
            if (!visited.Add(value)) return;
            if (value is PdfArray array) { foreach (var item in array.Items) Visit(item, depth + 1); return; }
            if (value is PdfStream stream) { Visit(stream.Dictionary, depth + 1); return; }
            if (value is not PdfDictionary dictionary) return;
            if (dictionary.Get<PdfName>("S")?.Name == "GoTo" && dictionary.Items.TryGetValue("SD", out var destination) &&
                PdfObjectLookup.Resolve(objects, destination) is PdfStringObj id && ids.TryGetValue(Convert.ToBase64String(id.RawBytes), out var mapped))
                dictionary.Items["SD"] = mapped;
            foreach (var child in dictionary.Items.Values) if (child is not PdfReference) Visit(child, depth + 1);
        }
        foreach (int number in numbers) Visit(objects[number].Value, 0);
    }

    internal static void PruneStructure(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog,
        HashSet<int> removedPages, HashSet<int> removedAnnotations, CancellationToken token) {
        if (!catalog.Items.TryGetValue("StructTreeRoot", out var value) || PdfObjectLookup.Resolve(objects, value) is not PdfDictionary root)
            throw new InvalidOperationException("The tagged document requires a readable structure-tree root.");
        var retainedElements = new HashSet<int>(); var visited = new HashSet<PdfDictionary>();
        if (root.Items.TryGetValue("K", out var kids)) root.Items["K"] = Prune(kids, null, 0) ?? new PdfArray();
        var removedKeys = new HashSet<int>();
        foreach (int page in removedPages) {
            if (((PdfDictionary)objects[page].Value).Get<PdfNumber>("StructParents") is PdfNumber key) removedKeys.Add(CheckedKey(key));
        }
        foreach (int annotation in removedAnnotations) {
            if (objects.TryGetValue(annotation, out var indirect) && indirect.Value is PdfDictionary dictionary &&
                dictionary.Get<PdfNumber>("StructParent") is PdfNumber key) removedKeys.Add(CheckedKey(key));
        }
        if (root.Items.TryGetValue("ParentTree", out var parentTree)) {
            var entries = ReadNumberTree(objects, parentTree, token); var nums = new PdfArray();
            foreach (var pair in entries) {
                if (removedKeys.Contains(pair.Key)) continue;
                var entry = PdfObjectLookup.Resolve(objects, pair.Value);
                if (entry is PdfArray parents) {
                    var filtered = new PdfArray();
                    foreach (var item in parents.Items) filtered.Items.Add(item is PdfReference reference && retainedElements.Contains(reference.ObjectNumber) ? item : PdfNull.Instance);
                    nums.Items.Add(new PdfNumber(pair.Key)); nums.Items.Add(filtered);
                } else if (pair.Value is PdfReference reference) {
                    if (retainedElements.Contains(reference.ObjectNumber)) { nums.Items.Add(new PdfNumber(pair.Key)); nums.Items.Add(pair.Value); }
                } else if (entry is not PdfNull) {
                    throw new InvalidOperationException("The tagged parent tree contains an unsupported structure-parent value.");
                }
            }
            var replacement = new PdfDictionary(); replacement.Items["Nums"] = nums; root.Items["ParentTree"] = replacement;
        }
        if (root.Items.TryGetValue("IDTree", out var ids)) {
            var pairs = new List<(PdfObject Key, PdfObject Value)>(); var seen = new HashSet<PdfDictionary>();
            CollectIds(ids, 0); var names = new PdfArray();
            foreach (var pair in pairs) { names.Items.Add(pair.Key); names.Items.Add(pair.Value); }
            var replacement = new PdfDictionary(); replacement.Items["Names"] = names; root.Items["IDTree"] = replacement;
            void CollectIds(PdfObject tree, int depth) {
                token.ThrowIfCancellationRequested();
                RequireStructureDepth(depth);
                if (PdfObjectLookup.Resolve(objects, tree) is not PdfDictionary dictionary || !seen.Add(dictionary))
                    throw new InvalidOperationException("The tagged ID tree is cyclic or unreadable.");
                if (dictionary.Items.TryGetValue("Names", out var values) && PdfObjectLookup.Resolve(objects, values) is PdfArray array) {
                    if (array.Items.Count % 2 != 0) throw new InvalidOperationException("The tagged ID tree has an incomplete entry.");
                    for (int index = 0; index < array.Items.Count; index += 2)
                        if (array.Items[index + 1] is PdfReference reference && retainedElements.Contains(reference.ObjectNumber))
                            pairs.Add((array.Items[index], array.Items[index + 1]));
                }
                if (dictionary.Items.TryGetValue("Kids", out var children) && PdfObjectLookup.Resolve(objects, children) is PdfArray childArray)
                    foreach (var child in childArray.Items) CollectIds(child, depth + 1);
            }
        }

        PdfObject? Prune(PdfObject item, int? inheritedPage, int depth) {
            token.ThrowIfCancellationRequested();
            RequireStructureDepth(depth);
            var resolved = PdfObjectLookup.Resolve(objects, item);
            if (resolved is PdfArray array) {
                var remaining = new PdfArray(); foreach (var child in array.Items) { var result = Prune(child, inheritedPage, depth + 1); if (result != null) remaining.Items.Add(result); }
                return remaining.Items.Count == 0 ? null : remaining;
            }
            if (resolved is PdfNumber) {
                if (!inheritedPage.HasValue) throw new InvalidOperationException("A tagged marked-content identifier has no page owner.");
                return removedPages.Contains(inheritedPage.Value) ? null : item;
            }
            if (resolved is not PdfDictionary dictionary || !visited.Add(dictionary))
                throw new InvalidOperationException("The tagged structure tree is cyclic, shared, or unreadable.");
            int? page = dictionary.Items.TryGetValue("Pg", out var pageValue) && pageValue is PdfReference owner ? owner.ObjectNumber : inheritedPage;
            string? type = dictionary.Get<PdfName>("Type")?.Name;
            if (type is "MCR" or "OBJR" || type == null && dictionary.Items.ContainsKey("MCID")) {
                if (!page.HasValue) throw new InvalidOperationException("A tagged content reference has no page owner.");
                if (removedPages.Contains(page.Value) || type == "OBJR" && dictionary.Items.TryGetValue("Obj", out var target) &&
                    target is PdfReference annotation && removedAnnotations.Contains(annotation.ObjectNumber)) return null;
                return item;
            }
            bool hadKids = dictionary.Items.TryGetValue("K", out var children);
            var remainingKids = hadKids ? Prune(children!, page, depth + 1) : null;
            if (hadKids && remainingKids is null || !hadKids && page.HasValue && removedPages.Contains(page.Value)) return null;
            if (remainingKids != null) dictionary.Items["K"] = remainingKids;
            if (dictionary.Items.TryGetValue("Pg", out var oldPage) && oldPage is PdfReference old && removedPages.Contains(old.ObjectNumber)) dictionary.Items.Remove("Pg");
            if (item is not PdfReference reference) throw new InvalidOperationException("Structure elements must be indirect objects.");
            retainedElements.Add(reference.ObjectNumber); return item;
        }
    }

    private static SortedDictionary<int, PdfObject> ReadNumberTree(Dictionary<int, PdfIndirectObject> objects, PdfObject value, CancellationToken token) {
        var result = new SortedDictionary<int, PdfObject>(); var visited = new HashSet<PdfDictionary>(); Visit(value, 0); return result;
        void Visit(PdfObject tree, int depth) {
            token.ThrowIfCancellationRequested();
            RequireStructureDepth(depth);
            if (PdfObjectLookup.Resolve(objects, tree) is not PdfDictionary dictionary || !visited.Add(dictionary))
                throw new InvalidOperationException("The tagged parent tree is cyclic or unreadable.");
            if (dictionary.Items.TryGetValue("Nums", out var nums) && PdfObjectLookup.Resolve(objects, nums) is PdfArray array) {
                if (array.Items.Count % 2 != 0) throw new InvalidOperationException("The tagged parent tree has an incomplete entry.");
                for (int index = 0; index < array.Items.Count; index += 2) {
                    if (array.Items[index] is not PdfNumber key) throw new InvalidOperationException("A tagged parent-tree key is not an integer.");
                    int number = CheckedKey(key);
                    if (result.ContainsKey(number)) throw new InvalidOperationException("The tagged parent tree has repeated keys.");
                    result.Add(number, array.Items[index + 1]);
                }
            }
            if (dictionary.Items.TryGetValue("Kids", out var kids) && PdfObjectLookup.Resolve(objects, kids) is PdfArray children)
                foreach (var child in children.Items) Visit(child, depth + 1);
        }
    }
    private static int CheckedKey(PdfNumber key) {
        if (key.Value < 0 || key.Value > int.MaxValue || Math.Floor(key.Value) != key.Value)
            throw new InvalidOperationException("A tagged parent-tree key is not a supported nonnegative integer.");
        return (int)key.Value;
    }

    private static void RebuildStructureIds(Dictionary<int, PdfIndirectObject> objects, PdfDictionary root, CancellationToken token) {
        var pairs = new List<(PdfStringObj Id, PdfReference Element)>(); var visited = new HashSet<int>();
        var activeArrays = new HashSet<PdfArray>();
        if (root.Items.TryGetValue("K", out var kids)) Visit(kids, 0);
        if (pairs.Count == 0) { root.Items.Remove("IDTree"); return; }
        pairs.Sort((left, right) => CompareBytes(left.Id.RawBytes, right.Id.RawBytes));
        var names = new PdfArray();
        for (int index = 0; index < pairs.Count; index++) {
            if (index > 0 && CompareBytes(pairs[index - 1].Id.RawBytes, pairs[index].Id.RawBytes) == 0)
                throw new InvalidOperationException("The tagged structure tree has repeated element IDs.");
            names.Items.Add(pairs[index].Id); names.Items.Add(pairs[index].Element);
        }
        var tree = new PdfDictionary(); tree.Items["Names"] = names; root.Items["IDTree"] = tree;
        void Visit(PdfObject value, int depth) {
            token.ThrowIfCancellationRequested();
            RequireStructureDepth(depth);
            var resolved = PdfObjectLookup.Resolve(objects, value);
            if (resolved is PdfArray array) {
                if (!activeArrays.Add(array)) throw new InvalidOperationException("The tagged structure tree contains a cyclic child array.");
                foreach (var item in array.Items) Visit(item, depth + 1);
                activeArrays.Remove(array);
                return;
            }
            if (resolved is not PdfDictionary dictionary || value is not PdfReference reference || !visited.Add(reference.ObjectNumber)) return;
            if (dictionary.Items.TryGetValue("ID", out var id) && PdfObjectLookup.Resolve(objects, id) is PdfStringObj text) pairs.Add((text, reference));
            if (dictionary.Items.TryGetValue("K", out var children)) Visit(children, depth + 1);
        }
        static int CompareBytes(byte[] left, byte[] right) {
            for (int index = 0; index < Math.Min(left.Length, right.Length); index++) { int order = left[index].CompareTo(right[index]); if (order != 0) return order; }
            return left.Length.CompareTo(right.Length);
        }
    }
}
