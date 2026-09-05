namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private const int MaximumSignatureFieldDepth = 128;

    internal static PdfUnsignedDerivativeResult CreateUnsignedDerivative(
        byte[] pdf,
        PdfLoadOptions? readOptions,
        System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNull(pdf, nameof(pdf));
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf, readOptions, cancellationToken);
        int signatureCount = Math.Max(preflight.Probe.Security.SignatureCount, preflight.Probe.HasSignatures ? 1 : 0);
        if (preflight.Probe.HasEncryption && !preflight.Probe.Security.HasOwnerAuthorization) {
            throw new InvalidOperationException("Creating an unsigned derivative of an encrypted PDF requires owner authorization.");
        }

        cancellationToken.ThrowIfCancellationRequested();
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions, out _, out _, cancellationToken);
        int catalogObjectNumber = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalogObjectNumber == 0 || objects[catalogObjectNumber].Value is not PdfDictionary catalog) {
            throw new InvalidDataException("PDF does not contain a readable catalog.");
        }

        var removedFieldObjectNumbers = new HashSet<int>();
        if (PdfObjectLookup.Resolve(objects, catalog.Items.TryGetValue("AcroForm", out PdfObject? acroFormObject) ? acroFormObject : null) is PdfDictionary acroForm) {
            if (PdfObjectLookup.Resolve(objects, acroForm.Items.TryGetValue("Fields", out PdfObject? fieldsObject) ? fieldsObject : null) is PdfArray fields) {
                RemoveSignatureFields(objects, fields, inheritedFieldType: null, removedFieldObjectNumbers, new HashSet<int>(), 0, cancellationToken);
                if (fields.Items.Count == 0) acroForm.Items.Remove("Fields");
            }
            RemoveReferences(acroForm, "CO", removedFieldObjectNumbers, objects, cancellationToken);
            acroForm.Items.Remove("SigFlags");
        }
        catalog.Items.Remove("Perms");
        catalog.Items.Remove("DSS");
        RemoveSignatureWidgetsFromPages(objects, catalog, removedFieldObjectNumbers, cancellationToken);
        PdfObjectGraphPruner.PruneUnreachableObjects(objects, catalogObjectNumber, cancellationToken);
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        byte[] rewritten = RewriteAllObjects(objects, catalogObjectNumber, document.UncheckedMetadata, pdf, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocumentPreflight result = PdfInspector.Preflight(rewritten, options: null, cancellationToken: cancellationToken);
        if (result.Probe.HasSignatures || result.Probe.HasEncryption) {
            throw new InvalidDataException("The unsigned derivative rewrite retained signature or encryption state.");
        }
        return new PdfUnsignedDerivativeResult(rewritten, signatureCount);
    }

    private static void RemoveReferences(
        PdfDictionary dictionary,
        string key,
        HashSet<int> removedObjectNumbers,
        Dictionary<int, PdfIndirectObject> objects,
        System.Threading.CancellationToken cancellationToken) {
        if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue(key, out PdfObject? value) ? value : null) is not PdfArray references) return;
        for (int index = references.Items.Count - 1; index >= 0; index--) {
            cancellationToken.ThrowIfCancellationRequested();
            if (references.Items[index] is PdfReference reference && removedObjectNumbers.Contains(reference.ObjectNumber)) references.Items.RemoveAt(index);
        }
        if (references.Items.Count == 0) dictionary.Items.Remove(key);
    }

    private static void RemoveSignatureFields(
        Dictionary<int, PdfIndirectObject> objects,
        PdfArray fields,
        string? inheritedFieldType,
        HashSet<int> removedFieldObjectNumbers,
        HashSet<int> visitedFieldObjectNumbers,
        int depth,
        System.Threading.CancellationToken cancellationToken) {
        if (depth > MaximumSignatureFieldDepth) throw new InvalidDataException("PDF signature field nesting exceeds the supported depth.");
        for (int index = fields.Items.Count - 1; index >= 0; index--) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject fieldObject = fields.Items[index];
            if (PdfObjectLookup.Resolve(objects, fieldObject) is not PdfDictionary field) continue;
            string? fieldType = TryReadName(objects, field, "FT") ?? inheritedFieldType;
            if (string.Equals(fieldType, "Sig", StringComparison.Ordinal)) {
                CollectFieldObjectNumbers(objects, fieldObject, removedFieldObjectNumbers, depth, cancellationToken);
                fields.Items.RemoveAt(index);
                continue;
            }
            if (fieldObject is PdfReference fieldReference && !visitedFieldObjectNumbers.Add(fieldReference.ObjectNumber)) {
                throw new InvalidDataException("PDF form field graph contains a cycle or reused field reference.");
            }
            if (PdfObjectLookup.Resolve(objects, field.Items.TryGetValue("Kids", out PdfObject? kidsObject) ? kidsObject : null) is PdfArray kids) {
                RemoveSignatureFields(objects, kids, fieldType, removedFieldObjectNumbers, visitedFieldObjectNumbers, depth + 1, cancellationToken);
                if (kids.Items.Count == 0) field.Items.Remove("Kids");
            }
        }
    }

    private static void CollectFieldObjectNumbers(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject fieldObject,
        HashSet<int> objectNumbers,
        int depth,
        System.Threading.CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (depth > MaximumSignatureFieldDepth) throw new InvalidDataException("PDF signature field nesting exceeds the supported depth.");
        if (fieldObject is PdfReference reference && !objectNumbers.Add(reference.ObjectNumber)) return;
        if (PdfObjectLookup.Resolve(objects, fieldObject) is not PdfDictionary field ||
            PdfObjectLookup.Resolve(objects, field.Items.TryGetValue("Kids", out PdfObject? kidsObject) ? kidsObject : null) is not PdfArray kids) return;
        foreach (PdfObject child in kids.Items) CollectFieldObjectNumbers(objects, child, objectNumbers, depth + 1, cancellationToken);
    }

    private static void RemoveSignatureWidgetsFromPages(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary catalog,
        HashSet<int> removedFieldObjectNumbers,
        System.Threading.CancellationToken cancellationToken) {
        if (!catalog.Items.TryGetValue("Pages", out PdfObject? pagesObject)) return;
        var visited = new HashSet<int>();
        VisitPageNode(pagesObject, 0);

        void VisitPageNode(PdfObject nodeObject, int depth) {
            cancellationToken.ThrowIfCancellationRequested();
            if (depth > MaximumSignatureFieldDepth) throw new InvalidDataException("PDF page tree nesting exceeds the supported depth while removing signature widgets.");
            if (nodeObject is PdfReference nodeReference && !visited.Add(nodeReference.ObjectNumber)) return;
            if (PdfObjectLookup.Resolve(objects, nodeObject) is not PdfDictionary node) return;
            if (string.Equals(TryReadName(objects, node, "Type"), "Page", StringComparison.Ordinal)) {
                if (PdfObjectLookup.Resolve(objects, node.Items.TryGetValue("Annots", out PdfObject? annotsObject) ? annotsObject : null) is not PdfArray annots) return;
                for (int index = annots.Items.Count - 1; index >= 0; index--) {
                    cancellationToken.ThrowIfCancellationRequested();
                    PdfObject annotationObject = annots.Items[index];
                    PdfDictionary? annotation = PdfObjectLookup.Resolve(objects, annotationObject) as PdfDictionary;
                    bool removedReference = annotationObject is PdfReference annotationReference && removedFieldObjectNumbers.Contains(annotationReference.ObjectNumber);
                    bool removedParent = annotation?.Items.TryGetValue("Parent", out PdfObject? parentObject) == true &&
                        parentObject is PdfReference parentReference && removedFieldObjectNumbers.Contains(parentReference.ObjectNumber);
                    if (removedReference || removedParent || annotation is not null && IsSignatureAnnotation(objects, annotation, cancellationToken)) annots.Items.RemoveAt(index);
                }
                if (annots.Items.Count == 0) node.Items.Remove("Annots");
                return;
            }
            if (PdfObjectLookup.Resolve(objects, node.Items.TryGetValue("Kids", out PdfObject? kidsObject) ? kidsObject : null) is PdfArray kids) {
                foreach (PdfObject child in kids.Items) VisitPageNode(child, depth + 1);
            }
        }
    }

    private static bool IsSignatureAnnotation(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary annotation,
        System.Threading.CancellationToken cancellationToken) {
        PdfDictionary? current = annotation;
        var visited = new HashSet<int>();
        for (int depth = 0; current is not null && depth <= MaximumSignatureFieldDepth; depth++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.Equals(TryReadName(objects, current, "FT"), "Sig", StringComparison.Ordinal)) return true;
            if (!current.Items.TryGetValue("Parent", out PdfObject? parentObject)) return false;
            if (parentObject is PdfReference parentReference && !visited.Add(parentReference.ObjectNumber)) {
                throw new InvalidDataException("PDF signature widget parent graph contains a cycle.");
            }
            current = PdfObjectLookup.Resolve(objects, parentObject) as PdfDictionary;
        }
        if (current is not null) throw new InvalidDataException("PDF signature widget parent nesting exceeds the supported depth.");
        return false;
    }
}
