using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Organizes tagged pages through the catalog-rooted writer without rebuilding the catalog from a subset.</summary>
internal static partial class PdfTaggedPageOrganizer {
    internal static bool IsTagged(PdfReadDocument document) =>
        document.Security.RootObjectNumber is int root && document.Objects[root].Value is PdfDictionary catalog &&
        catalog.Items.ContainsKey("StructTreeRoot");

    internal static byte[] Rewrite(byte[] source, PdfReadDocument document, IReadOnlyList<int> pageNumbers,
        PdfLoadOptions? options, long? maximumOutputBytes = null,
        Action<IReadOnlyDictionary<int, int>>? captureObjectNumbers = null, CancellationToken token = default) {
        Guard.NotNull(pageNumbers, nameof(pageNumbers)); token.ThrowIfCancellationRequested();
        if (pageNumbers.Count == 0 || pageNumbers.Any(page => page < 1 || page > document.Pages.Count))
            throw new ArgumentException("Select at least one existing page.", nameof(pageNumbers));
        var ordered = pageNumbers.Select(page => document.Pages[page - 1].ObjectNumber).ToArray();
        var retained = new HashSet<int>(ordered);
        var removed = new HashSet<int>(document.Pages.Select(page => page.ObjectNumber).Where(page => !retained.Contains(page)));
        // Compound extraction sessions reuse their source reader. Mutate a separately parsed graph for each output.
        return PdfDocumentObjectGraphRewriter.Rewrite(source, options, outputEncryption: null,
            mutateObjectGraph: (objects, security) => {
                var catalog = (PdfDictionary)objects[security.RootObjectNumber!.Value].Value;
                if (catalog.Items.TryGetValue("AcroForm", out var formValue) && PdfObjectLookup.Resolve(objects, formValue) is PdfDictionary form && form.Items.ContainsKey("XFA"))
                    throw new InvalidOperationException("Tagged page organization cannot update XFA page ownership.");
                var state = PdfPageExtractor.ExtractCatalogRewriteState(objects, document.TrailerRaw, document, token);
                var pruned = PdfPageExtractor.PruneCatalogStateForPages(objects, state, retained, ordered, cancellationToken: token);
                var overrides = PdfPageExtractor.BuildPageOverridesWithFilteredDestinationLinks(objects, ordered, null, pruned, retained, token);
                var removedAnnotations = new HashSet<int>();
                foreach (int page in removed) {
                    var dictionary = (PdfDictionary)objects[page].Value;
                    foreach (int annotation in AnnotationNumbers(objects, dictionary)) removedAnnotations.Add(annotation);
                }
                foreach (int page in ordered) {
                    token.ThrowIfCancellationRequested();
                    var dictionary = (PdfDictionary)objects[page].Value;
                    MaterializePageInheritance(objects, dictionary, token);
                    if (overrides != null && overrides.TryGetValue(page, out var values)) {
                        var previous = AnnotationNumbers(objects, dictionary).ToArray();
                        foreach (var value in values) dictionary.Items[value.Key] = value.Value;
                        var current = new HashSet<int>(AnnotationNumbers(objects, dictionary));
                        foreach (int annotation in previous.Where(number => !current.Contains(number))) removedAnnotations.Add(annotation);
                    }
                }
                ordered = CopyRepeatedPages(objects, catalog, ordered, token);
                ApplyNavigation(objects, catalog, pruned, removed, token);
                RemoveFieldsForAnnotations(objects, catalog, removedAnnotations, token);
                if (catalog.Items.ContainsKey("StructTreeRoot")) PruneStructure(objects, catalog, removed, removedAnnotations, token);
                int pagesObject = -1; while (objects.ContainsKey(pagesObject)) pagesObject--;
                var root = new PdfDictionary(); root.Items["Type"] = new PdfName("Pages"); root.Items["Count"] = new PdfNumber(ordered.Length);
                var kids = new PdfArray(); root.Items["Kids"] = kids;
                foreach (int page in ordered) {
                    kids.Items.Add(new PdfReference(page, objects[page].Generation));
                    ((PdfDictionary)objects[page].Value).Items["Parent"] = new PdfReference(pagesObject, 0);
                }
                objects.Add(pagesObject, new PdfIndirectObject(pagesObject, 0, root));
                catalog.Items["Pages"] = new PdfReference(pagesObject, 0);
                // Unknown catalog extensions may reference pages. Fail before serialization rather than retain a removed page invisibly.
                var collector = new PdfPageExtractor.ObjectCollector(objects, cancellationToken: token);
                collector.CollectObjectGraph(new PdfReference(security.RootObjectNumber.Value, objects[security.RootObjectNumber.Value].Generation));
                if (collector.ObjectIds.Any(removed.Contains))
                    throw new InvalidOperationException("Tagged page organization found an unsupported surviving reference to a removed page.");
                return security.InfoObjectNumber;
            }, maximumOutputBytes: maximumOutputBytes, cancellationToken: token, captureObjectNumbers: captureObjectNumbers);
    }

    private static IEnumerable<int> AnnotationNumbers(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page) {
        if (page.Items.TryGetValue("Annots", out var value) && PdfObjectLookup.Resolve(objects, value) is PdfArray annotations)
            foreach (var annotation in annotations.Items) if (annotation is PdfReference reference) yield return reference.ObjectNumber;
    }

    internal static void MaterializePageInheritance(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page, CancellationToken token) {
        foreach (string key in new[] { "Resources", "MediaBox", "CropBox", "Rotate" }) {
            var current = page; var visited = new HashSet<PdfDictionary>();
            while (visited.Add(current)) {
                token.ThrowIfCancellationRequested();
                if (current.Items.TryGetValue(key, out var value)) { page.Items[key] = value; break; }
                if (!current.Items.TryGetValue("Parent", out var parent) || PdfObjectLookup.Resolve(objects, parent) is not PdfDictionary ancestor) break;
                current = ancestor;
            }
        }
    }

    private static void RemoveFieldsForAnnotations(Dictionary<int, PdfIndirectObject> objects, PdfDictionary catalog,
        HashSet<int> removedAnnotations, CancellationToken token) {
        if (removedAnnotations.Count == 0 || !catalog.Items.TryGetValue("AcroForm", out var value) ||
            PdfObjectLookup.Resolve(objects, value) is not PdfDictionary form) return;
        if (form.Items.ContainsKey("XFA")) throw new InvalidOperationException("Tagged page organization cannot update XFA field ownership.");
        var removedFields = new HashSet<int>(); var visited = new HashSet<PdfDictionary>();
        if (form.Items.TryGetValue("Fields", out var fields) && PdfObjectLookup.Resolve(objects, fields) is PdfArray roots) PruneFields(roots);
        if (form.Items.TryGetValue("CO", out var order) && PdfObjectLookup.Resolve(objects, order) is PdfArray calculations)
            calculations.Items.RemoveAll(item => item is PdfReference reference && removedFields.Contains(reference.ObjectNumber));
        void PruneFields(PdfArray array) {
            for (int index = array.Items.Count - 1; index >= 0; index--) {
                token.ThrowIfCancellationRequested(); var item = array.Items[index];
                if (PdfObjectLookup.Resolve(objects, item) is not PdfDictionary field || !visited.Add(field))
                    throw new InvalidOperationException("Tagged page organization requires an unambiguous AcroForm field tree.");
                bool remove = item is PdfReference reference && removedAnnotations.Contains(reference.ObjectNumber);
                if (field.Items.TryGetValue("Kids", out var kids) && PdfObjectLookup.Resolve(objects, kids) is PdfArray children) {
                    PruneFields(children); remove |= children.Items.Count == 0;
                }
                if (remove) { if (item is PdfReference deleted) removedFields.Add(deleted.ObjectNumber); array.Items.RemoveAt(index); }
            }
        }
    }
}
