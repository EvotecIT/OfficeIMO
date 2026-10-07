namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationEditor {
    internal enum BatchOperation { Move, Copy, Group, Ungroup, Arrange, Remove }

    internal static PdfAnnotationEditResult EditBatch(byte[] pdf, IReadOnlyList<int> objectNumbers,
        BatchOperation operation, PdfLoadOptions? readOptions, double deltaX = 0D, double deltaY = 0D,
        PdfAnnotationOrderChange orderChange = PdfAnnotationOrderChange.Raise, bool allowResidualDataInAppendOnly = false) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(objectNumbers, nameof(objectNumbers));
        if (objectNumbers.Count == 0) throw new ArgumentException("Select at least one annotation.", nameof(objectNumbers));
        if (objectNumbers.Any(number => number <= 0)) throw new ArgumentOutOfRangeException(nameof(objectNumbers), "Annotation object numbers must be positive.");
        ValidateTransformFinite(deltaX, nameof(deltaX));
        ValidateTransformFinite(deltaY, nameof(deltaY));
        if (orderChange is not (PdfAnnotationOrderChange.Raise or PdfAnnotationOrderChange.Lower or PdfAnnotationOrderChange.BringToFront or PdfAnnotationOrderChange.SendToBack)) throw new ArgumentOutOfRangeException(nameof(orderChange));
        PdfMutationPlan plan = RequireAnnotationMutation(pdf, readOptions);
        if (operation == BatchOperation.Remove && plan.ExecutionMode == PdfMutationExecutionMode.AppendOnly && !allowResidualDataInAppendOnly)
            throw new NotSupportedException("Append-only annotation removal retains original data in older revisions. Explicitly permit residual data or use a permitted full rewrite.");
        PdfAnnotation[] selected = ResolveBatchAnnotations(GetAnnotationMutationDocumentInfo(plan).Annotations, objectNumbers);
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
        int catalog = FindCatalogObjectNumber(objects, trailerRaw);
        if (catalog == 0) throw new ArgumentException("PDF does not contain a readable catalog.", nameof(pdf));
        var changed = new HashSet<int>();
        HashSet<int>? removed = null;
        int additionalAnnotations = 0;
        switch (operation) {
            case BatchOperation.Move:
                foreach (PdfAnnotation annotation in selected) {
                    int number = annotation.ObjectNumber!.Value;
                    PdfAnnotationUpdateOptions options = CreateBatchMoveOptions(objects, annotation, deltaX, deltaY);
                    foreach (int generated in ApplyUpdates(objects, (PdfDictionary)objects[number].Value, options)) changed.Add(generated);
                    changed.Add(number);
                }
                break;
            case BatchOperation.Copy:
                additionalAnnotations = CopyBatchAnnotations(objects, selected, deltaX, deltaY, changed);
                break;
            case BatchOperation.Group:
                GroupBatchAnnotations(objects, selected, changed);
                break;
            case BatchOperation.Ungroup:
                foreach (PdfAnnotation annotation in selected.Where(annotation => annotation.Review?.IsGroup == true)) {
                    int number = annotation.ObjectNumber!.Value;
                    var dictionary = (PdfDictionary)objects[number].Value;
                    dictionary.Items.Remove("IRT"); dictionary.Items.Remove("RT"); changed.Add(number);
                }
                break;
            case BatchOperation.Arrange:
                ArrangeBatchAnnotations(objects, selected, orderChange, changed);
                break;
            case BatchOperation.Remove:
                removed = RemoveBatchAnnotations(objects, selected, changed);
                break;
        }
        if (changed.Count == 0) return new PdfAnnotationEditResult((byte[])pdf.Clone(), 0, plan, readOptions: readOptions);
        if (operation == BatchOperation.Group || operation == BatchOperation.Copy && selected.Any(annotation => annotation.Review?.IsGroup == true))
            EnsureGroupingVersion(pdf, objects, catalog, changed, plan.Preflight.Probe.Security);
        PdfGeneratedOutputGrowth growth = BuildGeneratedOutputGrowth(objects, changed,
            additionalAnnotationsPerPage: additionalAnnotations, additionalRevisions: plan.ExecutionMode == PdfMutationExecutionMode.AppendOnly ? 1 : 0);
        if (plan.ExecutionMode == PdfMutationExecutionMode.AppendOnly) {
            byte[] output = PdfIncrementalObjectWriter.Append(pdf, objects, plan.Preflight.Probe.Security, trailerRaw, changed,
                encryptionHandler: GetAppendEncryptionHandler(objects, trailerRaw, readOptions, plan.Preflight.Probe.Security));
            PdfLoadOptions outputOptions = PdfLoadOptions.ForGeneratedOutput(readOptions, pdf, output, growth);
            return new PdfAnnotationEditResult(output, selected.Length, plan, BuildAppendOnlyProof(pdf, output, plan, readOptions, outputOptions), readOptions: outputOptions);
        }
        if (removed is not null) {
            foreach (int number in removed) objects.Remove(number);
        }
        PdfObjectGraphPruner.PruneUnreachableObjects(objects, catalog);
        byte[] rewritten = RewriteAllObjects(objects, catalog, PdfReadDocument.Open(pdf, readOptions).UncheckedMetadata, pdf, out var map);
        return CreateFullRewriteResult(pdf, rewritten, selected.Length, plan, annotationsChanged: true, readOptions: readOptions,
            generatedGrowth: growth, objectNumberMap: map);
    }

    private static void EnsureGroupingVersion(byte[] pdf, Dictionary<int, PdfIndirectObject> objects, int catalogNumber, HashSet<int> changed,
        PdfDocumentSecurityInfo security) {
        if (PdfFileAssembler.ParseHeaderVersionOrDefault(PdfSyntax.GetHeaderVersion(pdf)) >= PdfFileVersion.Pdf16) return;
        var catalog = (PdfDictionary)objects[catalogNumber].Value;
        if (catalog.Items.TryGetValue("Version", out PdfObject? version) && PdfObjectLookup.Resolve(objects, version) is PdfName name &&
            PdfFileAssembler.ParseHeaderVersionOrDefault(name.Name) >= PdfFileVersion.Pdf16) return;
        if (security.HasDocMDPPermissions || security.HasUsageRights)
            throw new NotSupportedException("Grouping requires PDF 1.6. A certified or usage-rights document cannot receive the required catalog version upgrade.");
        // A catalog version upgrade is valid for PDF 1.4 without replacing input revision bytes.
        catalog.Items["Version"] = new PdfName("1.6");
        changed.Add(catalogNumber);
    }

    private static PdfAnnotation[] ResolveBatchAnnotations(IReadOnlyList<PdfAnnotation> annotations, IReadOnlyList<int> objectNumbers) {
        var selected = objectNumbers.Distinct().SelectMany(number => PdfAnnotationGrouping.GetMembers(annotations, number))
            .Distinct().ToArray();
        foreach (PdfAnnotation annotation in selected) {
            if (annotation.ObjectNumber is null || annotation.PageNumber is null ||
                !(IsAppearanceSubtype(annotation.Subtype) || annotation.Subtype is "Link" or "Redact"))
                throw new NotSupportedException("Batch editing supports indirect visual markup and link annotations attached to pages.");
            if (annotation.IsLocked || annotation.IsReadOnly) throw new InvalidOperationException("Locked or read-only annotations cannot be edited.");
            if (annotation.Review is { IsGroup: true } && PdfAnnotationGrouping.GetMembers(annotations, annotation.ObjectNumber.Value).Count < 2)
                throw new NotSupportedException("Malformed grouping relationships must be repaired before batch editing.");
        }
        return annotations.Where(selected.Contains).ToArray();
    }

    private static PdfAnnotationUpdateOptions CreateBatchMoveOptions(Dictionary<int, PdfIndirectObject> objects,
        PdfAnnotation annotation, double deltaX, double deltaY) => CreateTransformOptions(annotation,
        new PdfPageRectangle(annotation.X1 + deltaX, annotation.Y1 + deltaY, annotation.X2 + deltaX, annotation.Y2 + deltaY),
        annotation.Subtype == "Line" ? ReadLineAuxiliaryGeometry(objects, annotation.ObjectNumber!.Value) : default);

    private static void GroupBatchAnnotations(Dictionary<int, PdfIndirectObject> objects, PdfAnnotation[] selected, HashSet<int> changed) {
        if (selected.Length < 2 || selected.Select(annotation => annotation.PageNumber).Distinct().Count() != 1)
            throw new ArgumentException("A group requires at least two markup annotations on the same page.");
        if (selected.Any(annotation => annotation.Subtype is "Link" or "Redact" ||
            annotation.Review?.InReplyToObjectNumber is not null && annotation.Review.IsGroup != true))
            throw new NotSupportedException("Links, redaction marks, and conversational replies cannot become markup group members.");
        int primary = selected.First(annotation => annotation.Review?.InReplyToObjectNumber is null).ObjectNumber!.Value;
        foreach (PdfAnnotation annotation in selected) {
            int number = annotation.ObjectNumber!.Value;
            var dictionary = (PdfDictionary)objects[number].Value;
            if (number == primary) {
                if (dictionary.Items.Remove("IRT") | dictionary.Items.Remove("RT")) changed.Add(number);
            } else if (annotation.Review is not { IsGroup: true } review || review.InReplyToObjectNumber != primary) {
                dictionary.Items["IRT"] = new PdfReference(primary, objects[primary].Generation);
                dictionary.Items["RT"] = new PdfName("Group"); changed.Add(number);
            }
        }
    }
}
