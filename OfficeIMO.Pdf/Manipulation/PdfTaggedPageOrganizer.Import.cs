using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfTaggedPageOrganizer {
    /// <summary>Inserts an engine-created blank page without requiring permission to copy existing content.</summary>
    internal static byte[] InsertBlank(byte[] targetPdf, int insertBeforePage, double width, double height,
        PdfLoadOptions? targetOptions, CancellationToken token = default) {
        byte[] blank = PdfDocument.Create(compose => compose.Page(page => page.Size(width, height))).ToBytes(cancellationToken: token);
        return ImportCore(targetPdf, blank, insertBeforePage, [1], targetOptions, null,
            PdfMutationOperation.ModifyPageTree, token);
    }

    internal static byte[] Import(byte[] targetPdf, byte[] sourcePdf, int insertBeforePage,
        int[] selectedPages, PdfLoadOptions? targetOptions, PdfLoadOptions? sourceOptions, CancellationToken token = default) =>
        ImportCore(targetPdf, sourcePdf, insertBeforePage, selectedPages, targetOptions, sourceOptions, PdfMutationOperation.ImportPages, token);

    private static byte[] ImportCore(byte[] targetPdf, byte[] sourcePdf, int insertBeforePage,
        int[] selectedPages, PdfLoadOptions? targetOptions, PdfLoadOptions? sourceOptions,
        PdfMutationOperation targetOperation, CancellationToken token) {
        var (_, target) = PdfMutationPlanner.RequireFullRewriteDocument(targetPdf, targetOperation, targetOptions, cancellationToken: token);
        var (_, source) = PdfMutationPlanner.RequireFullRewriteDocument(sourcePdf, PdfMutationOperation.ImportPages, sourceOptions, cancellationToken: token);
        if (insertBeforePage < 1 || insertBeforePage > target.Pages.Count + 1) throw new ArgumentOutOfRangeException(nameof(insertBeforePage));
        int[] selection = selectedPages.Length == 0 ? Enumerable.Range(1, source.Pages.Count).ToArray() : selectedPages;
        if (selection.Length == 0) throw new ArgumentException("Source PDF does not contain any pages.", nameof(sourcePdf));
        if (selection.Any(page => page < 1 || page > source.Pages.Count))
            throw new ArgumentOutOfRangeException(nameof(selectedPages), "A selected page is outside the source document.");
        byte[] selected = Rewrite(sourcePdf, source, selection, sourceOptions, token: token);
        PdfLoadOptions selectedOptions = PdfLoadOptions.ForGeneratedOutput(sourceOptions, sourcePdf, selected);
        var incoming = PdfReadDocument.Open(selected, selectedOptions, token);
        var incomingCatalog = (PdfDictionary)incoming.Objects[incoming.Security.RootObjectNumber!.Value].Value;
        if (incoming.AcroFormXfa != null || incoming.FormFields.Count > 0)
            throw new NotSupportedException("Importing tagged page graphs with incoming form fields requires a form-name and resource collision policy.");
        if (incomingCatalog.Items.ContainsKey("OCProperties"))
            throw new NotSupportedException("Importing tagged page graphs with optional-content groups requires a layer collision policy.");
        if (incomingCatalog.Items.TryGetValue("Names", out var namesValue) && PdfObjectLookup.Resolve(incoming.Objects, namesValue) is PdfDictionary names &&
            names.Items.Keys.Any(key => key != "Dests"))
            throw new NotSupportedException("Importing tagged page graphs with document name trees other than destinations requires a catalog collision policy.");
        if (!IsTagged(incoming) && PdfInspector.FromReadDocument(target, PdfInspector.Probe(targetPdf), cancellationToken: token).XmpMetadata?.HasPdfUaIdentification == true)
            throw new NotSupportedException("Importing untagged pages into a PDF/UA document requires accessible tagging of the incoming content.");
        ResolveImportedDestinations(incoming, incomingCatalog, token);

        return PdfDocumentObjectGraphRewriter.Rewrite(targetPdf, targetOptions, null, (objects, security) => {
            var catalog = (PdfDictionary)objects[security.RootObjectNumber!.Value].Value;
            if (catalog.Items.TryGetValue("AcroForm", out var formValue) && PdfObjectLookup.Resolve(objects, formValue) is PdfDictionary form && form.Items.ContainsKey("XFA"))
                throw new NotSupportedException("Page insertion cannot update XFA page ownership.");
            int next = objects.Keys.Min() - 1;
            int Allocate() { while (objects.ContainsKey(next)) next--; return next--; }
            var map = new Dictionary<int, int>(); foreach (int number in incoming.Objects.Keys) map.Add(number, Allocate());
            PdfObject Clone(PdfObject value, int depth) {
                token.ThrowIfCancellationRequested();
                if (depth > 256) throw new InvalidOperationException("Page import exceeded the object nesting limit.");
                if (value is PdfReference reference) {
                    if (!map.TryGetValue(reference.ObjectNumber, out int number) || incoming.Objects[reference.ObjectNumber].Generation != reference.Generation)
                        throw new InvalidOperationException("Page import found an invalid source reference.");
                    return new PdfReference(number, 0);
                }
                if (value is PdfArray array) { var result = new PdfArray(); foreach (var item in array.Items) result.Items.Add(Clone(item, depth + 1)); return result; }
                if (value is PdfDictionary dictionary) { var result = new PdfDictionary(); foreach (var pair in dictionary.Items) result.Items[pair.Key] = Clone(pair.Value, depth + 1); return result; }
                if (value is PdfStream stream) return new PdfStream((PdfDictionary)Clone(stream.Dictionary, depth + 1), stream.Data, stream.DecodingFailed, stream.DecodingError);
                return value;
            }
            foreach (var pair in incoming.Objects) objects.Add(map[pair.Key], new PdfIndirectObject(map[pair.Key], 0, Clone(pair.Value.Value, 0)));
            var importedCatalog = (PdfDictionary)objects[map[incoming.Security.RootObjectNumber.Value]].Value;
            var originalPages = target.Pages.Select(page => page.ObjectNumber).ToArray();
            var importedPages = incoming.Pages.Select(page => map[page.ObjectNumber]).ToArray();
            foreach (int page in originalPages.Concat(importedPages)) MaterializePageInheritance(objects, (PdfDictionary)objects[page].Value, token);
            var ordered = originalPages.Take(insertBeforePage - 1).Concat(importedPages).Concat(originalPages.Skip(insertBeforePage - 1)).ToArray();
            MergeImportedStructure(objects, catalog, importedCatalog, map.Values, originalPages, insertBeforePage, token);
            int pagesNumber = Allocate(); var pages = new PdfDictionary(); pages.Items["Type"] = new PdfName("Pages"); pages.Items["Count"] = new PdfNumber(ordered.Length);
            var kids = new PdfArray(); pages.Items["Kids"] = kids;
            foreach (int page in ordered) { kids.Items.Add(new PdfReference(page, objects[page].Generation)); ((PdfDictionary)objects[page].Value).Items["Parent"] = new PdfReference(pagesNumber, 0); }
            objects.Add(pagesNumber, new PdfIndirectObject(pagesNumber, 0, pages)); catalog.Items["Pages"] = new PdfReference(pagesNumber, 0);
            var originalIndexes = originalPages.Select((page, index) => (page, index)).ToDictionary(item => item.page,
                item => item.index + (item.index >= insertBeforePage - 1 ? importedPages.Length : 0));
            var targetState = PdfPageExtractor.ExtractCatalogRewriteState(target.Objects, target.TrailerRaw, target, token);
            var labels = PdfPageExtractor.PruneCatalogStateForPages(objects, targetState, new HashSet<int>(originalPages), originalPages,
                outputPageIndexByPageObjectNumber: originalIndexes, cancellationToken: token).PageLabels;
            if (labels != null) catalog.Items["PageLabels"] = labels; else catalog.Items.Remove("PageLabels");
            return security.InfoObjectNumber;
        }, cancellationToken: token);
    }

    private static void ResolveImportedDestinations(PdfReadDocument source, PdfDictionary catalog, CancellationToken token) {
        var named = new Dictionary<string, PdfObject>(StringComparer.Ordinal);
        if (catalog.Items.TryGetValue("Dests", out var directValue) && PdfObjectLookup.Resolve(source.Objects, directValue) is PdfDictionary direct)
            foreach (var pair in direct.Items) named[pair.Key] = pair.Value;
        if (catalog.Items.TryGetValue("Names", out var namesValue) && PdfObjectLookup.Resolve(source.Objects, namesValue) is PdfDictionary names && names.Items.TryGetValue("Dests", out var tree)) Collect(tree, new HashSet<PdfDictionary>());
        void Collect(PdfObject value, HashSet<PdfDictionary> seen) {
            token.ThrowIfCancellationRequested();
            if (PdfObjectLookup.Resolve(source.Objects, value) is not PdfDictionary dictionary || !seen.Add(dictionary)) throw new InvalidOperationException("The source destination name tree is unreadable.");
            if (dictionary.Items.TryGetValue("Names", out var pairsValue) && PdfObjectLookup.Resolve(source.Objects, pairsValue) is PdfArray pairs)
                for (int index = 0; index + 1 < pairs.Items.Count; index += 2)
                    if (pairs.Items[index] is PdfStringObj name) named[name.Value] = pairs.Items[index + 1];
            if (dictionary.Items.TryGetValue("Kids", out var kidsValue) && PdfObjectLookup.Resolve(source.Objects, kidsValue) is PdfArray kids)
                foreach (var child in kids.Items) Collect(child, seen);
        }
        PdfObject Resolve(PdfObject value) {
            var seen = new HashSet<PdfObject>();
            while (seen.Add(value)) {
                token.ThrowIfCancellationRequested(); var resolved = PdfObjectLookup.Resolve(source.Objects, value);
                if (resolved is PdfDictionary dictionary && dictionary.Items.TryGetValue("D", out value!)) continue;
                if (resolved is PdfArray) return resolved;
                string? name = resolved is PdfName pdfName ? pdfName.Name : (resolved as PdfStringObj)?.Value;
                if (name == null || !named.TryGetValue(name, out value!)) throw new InvalidOperationException("An imported page destination cannot be resolved.");
            }
            throw new InvalidOperationException("The imported destination is cyclic.");
        }
        var visited = new HashSet<PdfObject>();
        void Visit(PdfObject value) {
            token.ThrowIfCancellationRequested(); if (!visited.Add(value)) return;
            if (value is PdfArray array) { foreach (var item in array.Items) Visit(item); return; }
            if (value is PdfStream stream) { Visit(stream.Dictionary); return; }
            if (value is not PdfDictionary dictionary) return;
            if (dictionary.Items.TryGetValue("Dest", out var destination)) dictionary.Items["Dest"] = Resolve(destination);
            if (dictionary.Get<PdfName>("S")?.Name == "GoTo" && dictionary.Items.TryGetValue("D", out destination)) dictionary.Items["D"] = Resolve(destination);
            foreach (var child in dictionary.Items.Values) if (child is not PdfReference) Visit(child);
        }
        foreach (var indirect in source.Objects.Values) Visit(indirect.Value);
    }
}
