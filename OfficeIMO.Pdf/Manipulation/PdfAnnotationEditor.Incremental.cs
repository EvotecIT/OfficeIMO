namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationEditor {
    private static PdfAnnotationEditResult UpdateAnnotationIncrementally(
        byte[] pdf,
        int objectNumber,
        PdfAnnotationUpdateOptions options,
        PdfMutationPlan mutationPlan,
        PdfReadOptions? readOptions) {
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        if (!objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect) ||
            indirect.Value is not PdfDictionary annotation ||
            !IsAnnotationUpdateTarget(objects, objectNumber, annotation)) {
            throw new ArgumentException("PDF annotation object was not found: " + objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".", nameof(objectNumber));
        }

        ThrowIfAppendOnlyWidgetMutation(objects, annotation);
        IReadOnlyList<int> additionalChangedObjects = ApplyUpdates(objects, annotation, options);
        byte[] updated = PdfIncrementalObjectWriter.Append(
            pdf,
            objects,
            mutationPlan.Preflight.Probe.Security,
            trailerRaw,
            new[] { objectNumber }.Concat(additionalChangedObjects).Distinct().ToArray(),
            encryptionHandler: GetAppendEncryptionHandler(objects, trailerRaw, readOptions, mutationPlan.Preflight.Probe.Security));
        PdfSignatureMutationReport proof = BuildAppendOnlyProof(pdf, updated, mutationPlan, readOptions);
        return new PdfAnnotationEditResult(updated, 1, mutationPlan, proof, readOptions: readOptions);
    }

    private static PdfAnnotationEditResult RemoveAnnotationsIncrementally(
        byte[] pdf,
        PdfAnnotationRemovalOptions options,
        PdfMutationPlan mutationPlan,
        PdfReadOptions? readOptions) {
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        List<int> pageObjectNumbers = GetPageObjectNumbersInDocumentOrder(objects);
        var removedObjectNumbers = new HashSet<int>();
        var changedObjectNumbers = new HashSet<int>();
        int directRemovalCount = 0;

        for (int pageIndex = 0; pageIndex < pageObjectNumbers.Count; pageIndex++) {
            if (options.PageNumber.HasValue && options.PageNumber.Value != pageIndex + 1) {
                continue;
            }

            int pageObjectNumber = pageObjectNumbers[pageIndex];
            if (!objects.TryGetValue(pageObjectNumber, out PdfIndirectObject? pageObject) ||
                pageObject.Value is not PdfDictionary page ||
                !page.Items.TryGetValue("Annots", out PdfObject? annotsObject) ||
                PdfObjectLookup.Resolve(objects, annotsObject) is not PdfArray annotations) {
                continue;
            }

            int annotationsOwnerObjectNumber = annotsObject is PdfReference annotationsReference
                ? annotationsReference.ObjectNumber
                : pageObjectNumber;
            bool annotationsChanged = false;
            for (int i = annotations.Items.Count - 1; i >= 0; i--) {
                PdfObject item = annotations.Items[i];
                int? annotationObjectNumber = item is PdfReference reference ? reference.ObjectNumber : null;
                PdfDictionary? annotation = PdfObjectLookup.Resolve(objects, item) as PdfDictionary;
                if (annotation is null || !MatchesRemovalFilter(objects, annotation, annotationObjectNumber, options)) {
                    continue;
                }

                ThrowIfAppendOnlyWidgetMutation(objects, annotation);
                if (options.RemoveMatchingPopups &&
                    annotation.Items.TryGetValue("Popup", out PdfObject? popupObject) &&
                    popupObject is PdfReference popupReference) {
                    removedObjectNumbers.Add(popupReference.ObjectNumber);
                }

                if (annotationObjectNumber.HasValue) {
                    removedObjectNumbers.Add(annotationObjectNumber.Value);
                } else {
                    directRemovalCount++;
                }

                annotations.Items.RemoveAt(i);
                annotationsChanged = true;
            }

            if (options.RemoveMatchingPopups) {
                RemoveIncrementalPopupReferences(objects, annotations, removedObjectNumbers, changedObjectNumbers, ref annotationsChanged);
            }

            if (annotations.Items.Count == 0 && annotsObject is not PdfReference) {
                page.Items.Remove("Annots");
                changedObjectNumbers.Add(pageObjectNumber);
            } else if (annotationsChanged) {
                changedObjectNumbers.Add(annotationsOwnerObjectNumber);
            }
        }

        int affectedCount = removedObjectNumbers.Count + directRemovalCount;
        if (changedObjectNumbers.Count == 0) {
            return new PdfAnnotationEditResult((byte[])pdf.Clone(), 0, mutationPlan, readOptions: readOptions);
        }

        byte[] updated = PdfIncrementalObjectWriter.Append(
            pdf,
            objects,
            mutationPlan.Preflight.Probe.Security,
            trailerRaw,
            changedObjectNumbers,
            encryptionHandler: GetAppendEncryptionHandler(objects, trailerRaw, readOptions, mutationPlan.Preflight.Probe.Security));
        PdfSignatureMutationReport proof = BuildAppendOnlyProof(pdf, updated, mutationPlan, readOptions);
        return new PdfAnnotationEditResult(updated, Math.Max(affectedCount, 1), mutationPlan, proof, readOptions: readOptions);
    }

    private static void RemoveIncrementalPopupReferences(
        Dictionary<int, PdfIndirectObject> objects,
        PdfArray annotations,
        HashSet<int> removedObjectNumbers,
        HashSet<int> changedObjectNumbers,
        ref bool annotationsChanged) {
        if (removedObjectNumbers.Count == 0) {
            return;
        }

        for (int i = annotations.Items.Count - 1; i >= 0; i--) {
            PdfObject item = annotations.Items[i];
            if (item is PdfReference reference && removedObjectNumbers.Contains(reference.ObjectNumber)) {
                annotations.Items.RemoveAt(i);
                annotationsChanged = true;
                continue;
            }

            if (PdfObjectLookup.Resolve(objects, item) is not PdfDictionary annotation ||
                !annotation.Items.TryGetValue("Popup", out PdfObject? popupObject) ||
                popupObject is not PdfReference popupReference ||
                !removedObjectNumbers.Contains(popupReference.ObjectNumber)) {
                continue;
            }

            annotation.Items.Remove("Popup");
            if (item is PdfReference annotationReference) {
                changedObjectNumbers.Add(annotationReference.ObjectNumber);
            } else {
                annotationsChanged = true;
            }
        }
    }

    private static void ThrowIfAppendOnlyWidgetMutation(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary annotation) {
        if (string.Equals(TryReadName(objects, annotation, "Subtype"), "Widget", StringComparison.OrdinalIgnoreCase)) {
            throw new NotSupportedException("Append-only widget annotation changes must use the form-field mutation engine so FieldMDP locks and appearances are evaluated.");
        }
    }

    private static PdfStandardSecurityHandler? GetAppendEncryptionHandler(
        Dictionary<int, PdfIndirectObject> objects,
        string trailerRaw,
        PdfReadOptions? readOptions,
        PdfDocumentSecurityInfo security) {
        if (!security.HasEncryption) return null;
        if (!PdfSyntax.TryCreateDecryptor(objects, trailerRaw, readOptions, out PdfStandardSecurityHandler? handler)) {
            throw new InvalidOperationException("Encrypted append-only annotation updates require authenticated PDF read options.");
        }

        return handler;
    }

    private static PdfSignatureMutationReport BuildAppendOnlyProof(
        byte[] before,
        byte[] after,
        PdfMutationPlan mutationPlan,
        PdfReadOptions? readOptions = null) {
        PdfSignatureMutationReport proof = PdfSignatureMutationAnalyzer.Analyze(
            before,
            after,
            PdfMutationOperation.ModifyAnnotations,
            readOptions: readOptions,
            executionPreference: mutationPlan.ExecutionPreference);
        if (!proof.IsPreservedAppendOnlyMutation) {
            throw new InvalidOperationException("Append-only annotation mutation did not preserve the input prefix, revision chain, and existing signature byte ranges.");
        }

        return proof;
    }

    private static PdfAnnotationEditResult CreateFullRewriteResult(
        byte[] source,
        byte[] rewritten,
        int affectedAnnotationCount,
        PdfMutationPlan mutationPlan,
        bool annotationsChanged,
        PdfReadOptions? readOptions) {
        PdfReadOptions rewrittenReadOptions = PdfReadOptions.WithMinimumInputBytes(readOptions, rewritten.LongLength);
        var preservationOptions = new PdfRewritePreservationOptions {
            OriginalReadOptions = readOptions,
            RewrittenReadOptions = rewrittenReadOptions,
            PreserveAnnotations = !annotationsChanged,
            PreserveLinkAnnotations = !annotationsChanged,
            PreserveRevisionStructure = false
        };
        PdfRewritePreservationReport preservation = PdfRewritePreservation.Assess(source, rewritten, preservationOptions);
        return new PdfAnnotationEditResult(rewritten, affectedAnnotationCount, mutationPlan, rewritePreservationReport: preservation, readOptions: readOptions);
    }
}
