using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfTaggedPageOrganizer {
    private static int[] CopyRepeatedPages(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog,
        int[] ordered, CancellationToken token) {
        var seen = new HashSet<int>(); var output = new int[ordered.Length];
        int next = objects.Keys.Min() - 1;
        for (int index = 0; index < ordered.Length; index++) {
            int page = ordered[index];
            if (seen.Add(page)) { output[index] = page; continue; }
            var sourcePage = (PdfDictionary)objects[page].Value;
            var copies = new Dictionary<int, int>();
            int Allocate() { while (objects.ContainsKey(next)) next--; return next--; }
            int copiedPage = Allocate(); copies.Add(page, copiedPage);
            var copy = (PdfDictionary)Clone(sourcePage, 0); objects.Add(copiedPage, new PdfIndirectObject(copiedPage, 0, copy));
            output[index] = copiedPage;
            if (copy.Items.TryGetValue("Annots", out var annotations) && PdfObjectLookup.Resolve(objects, annotations) is PdfArray annotationArray)
                foreach (var annotation in annotationArray.Items)
                    if (PdfObjectLookup.Resolve(objects, annotation) is PdfDictionary dictionary)
                        dictionary.Items["P"] = new PdfReference(copiedPage, 0);
            CopyStructure(objects, catalog, page, copiedPage, copies, Clone, Allocate, token);
            CopyFields(objects, catalog, copies, copiedPage, token);

            PdfObject Clone(PdfObject value, int depth) {
                token.ThrowIfCancellationRequested();
                if (depth > 256) throw new InvalidOperationException("Page copying exceeded the object nesting limit.");
                if (value is PdfReference reference) {
                    if (copies.TryGetValue(reference.ObjectNumber, out int existing)) return new PdfReference(existing, 0);
                    if (!objects.TryGetValue(reference.ObjectNumber, out var original) || original.Generation != reference.Generation)
                        throw new InvalidOperationException("Page copying found an invalid object reference.");
                    string? type = (original.Value as PdfDictionary)?.Get<PdfName>("Type")?.Name;
                    if (type is "Page" or "Pages" or "StructElem" or "StructTreeRoot" or "Catalog") return value;
                    int number = Allocate(); copies.Add(reference.ObjectNumber, number);
                    objects.Add(number, new PdfIndirectObject(number, 0, Clone(original.Value, depth + 1)));
                    return new PdfReference(number, 0);
                }
                if (value is PdfArray array) { var result = new PdfArray(); foreach (var item in array.Items) result.Items.Add(Clone(item, depth + 1)); return result; }
                if (value is PdfDictionary dictionary) { var result = new PdfDictionary(); foreach (var pair in dictionary.Items) result.Items[pair.Key] = Clone(pair.Value, depth + 1); return result; }
                if (value is PdfStream stream) return new PdfStream((PdfDictionary)Clone(stream.Dictionary, depth + 1), stream.Data, stream.DecodingFailed, stream.DecodingError);
                return value;
            }
        }
        return output;
    }

    private static void CopyStructure(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog, int sourcePage, int copiedPage,
        Dictionary<int, int> copies, Func<PdfObject, int, PdfObject> clone, Func<int> allocate, CancellationToken token) {
        if (!catalog.Items.TryGetValue("StructTreeRoot", out var rootValue) || PdfObjectLookup.Resolve(objects, rootValue) is not PdfDictionary root) return;
        var projected = new Dictionary<int, int>(); var visited = new HashSet<int>();
        var ids = new Dictionary<string, PdfStringObj>(StringComparer.Ordinal);
        string idPrefix = UniqueStructureIdPrefix(objects, "copy" + (-copiedPage).ToString(System.Globalization.CultureInfo.InvariantCulture) + ":", token);
        PdfObject? Project(PdfObject value, int? inheritedPage, PdfObject parent, int depth) {
            token.ThrowIfCancellationRequested();
            if (depth > 256) throw new InvalidOperationException("Page copying exceeded the structure nesting limit.");
            var resolved = PdfObjectLookup.Resolve(objects, value);
            if (resolved is PdfArray array) { var result = new PdfArray(); foreach (var item in array.Items) { var child = Project(item, inheritedPage, parent, depth + 1); if (child != null) result.Items.Add(child); } return result.Items.Count == 0 ? null : result; }
            if (resolved is PdfNumber) {
                if (!inheritedPage.HasValue) throw new InvalidOperationException("A marked-content identifier has no page owner.");
                return inheritedPage == sourcePage ? value : null;
            }
            if (resolved is not PdfDictionary dictionary) throw new InvalidOperationException("Page copying found an unreadable structure child.");
            int? page = dictionary.Items.TryGetValue("Pg", out var owner) && owner is PdfReference pageRef ? pageRef.ObjectNumber : inheritedPage;
            string? type = dictionary.Get<PdfName>("Type")?.Name;
            if (type is "MCR" or "OBJR" || type == null && dictionary.Items.ContainsKey("MCID")) {
                if (!page.HasValue) throw new InvalidOperationException("A content reference has no page owner.");
                if (page != sourcePage) return null;
                var result = (PdfDictionary)clone(dictionary, 0); result.Items["Pg"] = new PdfReference(copiedPage, 0); return result;
            }
            if (value is not PdfReference reference || !visited.Add(reference.ObjectNumber)) throw new InvalidOperationException("Page copying requires an unambiguous structure tree.");
            int number = allocate(); var resultReference = new PdfReference(number, 0);
            var kids = dictionary.Items.TryGetValue("K", out var k) ? Project(k, page, resultReference, depth + 1) : null;
            if (kids == null && page != sourcePage) return null;
            var resultDictionary = new PdfDictionary();
            foreach (var pair in dictionary.Items)
                if (pair.Key is not ("P" or "Pg" or "K" or "ID")) resultDictionary.Items[pair.Key] = clone(pair.Value, 0);
            resultDictionary.Items["P"] = parent;
            if (page == sourcePage) resultDictionary.Items["Pg"] = new PdfReference(copiedPage, 0);
            if (kids != null) resultDictionary.Items["K"] = kids;
            if (dictionary.Items.TryGetValue("ID", out var id) && PdfObjectLookup.Resolve(objects, id) is PdfStringObj text) {
                var mapped = PrefixStructureId(idPrefix, text);
                resultDictionary.Items["ID"] = mapped; ids.Add(Convert.ToBase64String(text.RawBytes), mapped);
            }
            objects.Add(number, new PdfIndirectObject(number, 0, resultDictionary)); projected.Add(reference.ObjectNumber, number); return resultReference;
        }
        PdfDictionary parentDictionary = root; PdfObject parentValue = rootValue; PdfObject? originalKids = root.Items.TryGetValue("K", out var rootKids) ? rootKids : null;
        var roots = PdfObjectLookup.Resolve(objects, originalKids) is PdfArray rootArray ? rootArray.Items.ToArray() : originalKids == null ? Array.Empty<PdfObject>() : new[] { originalKids };
        if (roots.Length == 1 && PdfObjectLookup.Resolve(objects, roots[0]) is PdfDictionary document && document.Get<PdfName>("S")?.Name == "Document") {
            parentDictionary = document; parentValue = roots[0]; originalKids = document.Items.TryGetValue("K", out var kids) ? kids : null;
        }
        var combined = new PdfArray(); var original = PdfObjectLookup.Resolve(objects, originalKids);
        var originalChildren = original is PdfArray array ? array.Items.ToArray() : originalKids == null ? Array.Empty<PdfObject>() : new[] { originalKids };
        bool copyDocument = parentDictionary != root && originalChildren.Any(child =>
            PdfObjectLookup.Resolve(objects, child) is not PdfDictionary element || element.Get<PdfName>("Type")?.Name is "MCR" or "OBJR" ||
            element.Items.ContainsKey("MCID"));
        foreach (var child in originalChildren) {
            combined.Items.Add(child);
            if (copyDocument) continue;
            var added = Project(child, parentDictionary.Get<PdfReference>("Pg")?.ObjectNumber, parentValue, 0);
            if (added is PdfArray addedArray) combined.Items.AddRange(addedArray.Items); else if (added != null) combined.Items.Add(added);
        }
        // A Document may own marked content directly. Its copied content needs a
        // separate page owner, while the original Document remains the sole root.
        if (copyDocument && Project(parentValue, null, parentValue, 0) is PdfReference copiedDocument) {
            ((PdfDictionary)objects[copiedDocument.ObjectNumber].Value).Items["S"] = new PdfName("Part");
            combined.Items.Add(copiedDocument);
        }
        parentDictionary.Items["K"] = combined;
        RemapStructureDestinations(objects, copies.Values, ids, token);
        var entries = root.Items.TryGetValue("ParentTree", out var parentTree) ? ReadNumberTree(objects, parentTree, token) : new SortedDictionary<int, PdfObject>();
        int nextKey = entries.Count == 0 ? 0 : checked(entries.Keys.Max() + 1);
        foreach (var pair in copies.ToArray()) {
            var old = objects[pair.Key].Value is PdfStream oldStream ? oldStream.Dictionary : objects[pair.Key].Value as PdfDictionary;
            var current = objects[pair.Value].Value is PdfStream stream ? stream.Dictionary : objects[pair.Value].Value as PdfDictionary;
            if (old == null || current == null) continue;
            foreach (string key in new[] { "StructParents", "StructParent" }) if (old.Get<PdfNumber>(key) is PdfNumber number) {
                int originalKey = CheckedKey(number);
                if (!entries.TryGetValue(originalKey, out var value)) throw new InvalidOperationException("A copied tagged object has no parent-tree entry.");
                PdfObject MapParent(PdfObject item) {
                    if (item is PdfNull) return item;
                    if (item is PdfReference reference && projected.TryGetValue(reference.ObjectNumber, out int mapped)) return new PdfReference(mapped, 0);
                    throw new InvalidOperationException("A copied tagged object's parent belongs to another page or is not in the structure tree.");
                }
                var resolved = PdfObjectLookup.Resolve(objects, value);
                PdfObject mappedValue;
                if (resolved is PdfArray parents) { var mapped = new PdfArray(); foreach (var item in parents.Items) mapped.Items.Add(MapParent(item)); mappedValue = mapped; }
                else mappedValue = MapParent(value);
                current.Items[key] = new PdfNumber(nextKey); entries.Add(nextKey++, mappedValue);
            }
        }
        var nums = new PdfArray(); foreach (var pair in entries) { nums.Items.Add(new PdfNumber(pair.Key)); nums.Items.Add(pair.Value); }
        var replacement = new PdfDictionary(); replacement.Items["Nums"] = nums; root.Items["ParentTree"] = replacement;
        root.Items["ParentTreeNextKey"] = new PdfNumber(nextKey);
        if (projected.Count > 0) RebuildStructureIds(objects, root, token);
    }

    private static void CopyFields(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog, Dictionary<int, int> copies,
        int copiedPage, CancellationToken token) {
        if (!catalog.Items.TryGetValue("AcroForm", out var value) || PdfObjectLookup.Resolve(objects, value) is not PdfDictionary form ||
            !form.Items.TryGetValue("Fields", out var fieldsValue) || PdfObjectLookup.Resolve(objects, fieldsValue) is not PdfArray fields) return;
        var names = new HashSet<string>(fields.Items.Select(item => (PdfObjectLookup.Resolve(objects, item) as PdfDictionary)?.Get<PdfStringObj>("T")?.Value ?? string.Empty), StringComparer.Ordinal);
        foreach (var original in fields.Items.ToArray()) {
            if (original is not PdfReference root || !copies.TryGetValue(root.ObjectNumber, out int number)) continue;
            var reference = new PdfReference(number, 0);
            if (!Retain(reference, new HashSet<int>())) continue;
            var dictionary = (PdfDictionary)objects[number].Value;
            if (dictionary.Get<PdfStringObj>("T") is not PdfStringObj name) throw new InvalidOperationException("Copying an unnamed form field is not supported.");
            string copiedName = name.Value + "_copy" + (-copiedPage).ToString(System.Globalization.CultureInfo.InvariantCulture);
            while (!names.Add(copiedName)) copiedName += "_";
            dictionary.Items["T"] = new PdfStringObj(copiedName, true);
            fields.Items.Add(reference);
            if (form.Items.TryGetValue("CO", out var order) && PdfObjectLookup.Resolve(objects, order) is PdfArray calculations) {
                var retained = new HashSet<int>();
                Collect(reference);
                foreach (var item in calculations.Items.ToArray())
                    if (item is PdfReference calculation && copies.TryGetValue(calculation.ObjectNumber, out int mapped) && retained.Contains(mapped))
                        calculations.Items.Add(new PdfReference(mapped, objects[mapped].Generation));
                void Collect(PdfObject item) {
                    token.ThrowIfCancellationRequested();
                    if (item is not PdfReference field || !retained.Add(field.ObjectNumber) ||
                        PdfObjectLookup.Resolve(objects, field) is not PdfDictionary definition) return;
                    if (definition.Items.TryGetValue("Kids", out var kids) && PdfObjectLookup.Resolve(objects, kids) is PdfArray children)
                        foreach (var child in children.Items) Collect(child);
                }
            }
        }
        bool Retain(PdfObject item, HashSet<int> visited) {
            token.ThrowIfCancellationRequested();
            if (item is not PdfReference reference || !visited.Add(reference.ObjectNumber) || PdfObjectLookup.Resolve(objects, item) is not PdfDictionary dictionary)
                throw new InvalidOperationException("Copying requires an unambiguous AcroForm field tree.");
            if (dictionary.Items.ContainsKey("AA")) throw new InvalidOperationException("Copying form fields with automatic actions requires an explicit field-name policy.");
            if (dictionary.Get<PdfName>("Subtype")?.Name == "Widget" && dictionary.Get<PdfReference>("P")?.ObjectNumber != copiedPage) return false;
            if (dictionary.Items.TryGetValue("Kids", out var kidsValue) && PdfObjectLookup.Resolve(objects, kidsValue) is PdfArray kids) {
                kids.Items.RemoveAll(child => !Retain(child, visited)); return kids.Items.Count > 0;
            }
            return true;
        }
    }
}
