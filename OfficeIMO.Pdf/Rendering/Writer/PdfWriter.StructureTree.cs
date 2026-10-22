using System.Text;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static void BuildGeneratedStructTree(IList<byte[]> objects, IReadOnlyList<LayoutResult.Page> pages, List<int> pageIds, int structTreeRootId, string? documentLanguage, bool pdf20 = false) {
        if (!pdf20 && !pages.Any(page => page.StructElements.Count > 0)) {
            ReplaceObject(objects, structTreeRootId, PdfStructTreeRootDictionaryBuilder.BuildEmptyStructTreeRootDictionary());
            return;
        }

        int documentStructElementId = ReserveObject(objects);
        var documentChildren = new List<PageStructElement>();
        var parentTreeEntries = new List<PdfStructTreeRootDictionaryBuilder.ParentTreeEntry>();
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            LayoutResult.Page page = pages[pageIndex];
            for (int elementIndex = 0; elementIndex < page.StructElements.Count; elementIndex++) {
                page.StructElements[elementIndex].ObjectId = ReserveObject(objects);
            }
        }

        int? namespaceId = pdf20 ? AddObject(objects, "<< /Type /Namespace /NS (http://iso.org/pdf2/ssn) >>\n") : null;
        var structureParents = BuildStructureParents(pages);
        var tableHeaders = BuildTableHeaderAssociations(pages, structureParents);
        var structureChildren = new StructureChildIndex(pages);
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            LayoutResult.Page page = pages[pageIndex];
            if (page.StructElements.Count == 0) {
                continue;
            }

            for (int elementIndex = 0; elementIndex < page.StructElements.Count; elementIndex++) {
                PageStructElement element = page.StructElements[elementIndex];
                int? contentStreamObjectId = element.MarkedContentId.HasValue
                    ? FindMarkedContentStreamObjectId(page, element.MarkedContentId.Value)
                    : null;
                List<int?>? additionalContentStreamObjectIds = element.AdditionalMarkedContentIds?
                    .Select(id => FindMarkedContentStreamObjectId(page, id))
                    .ToList();
                int parentObjectId = element.ParentElement != null
                    ? element.ParentElement.ObjectId
                    : element.ParentElementIndex.HasValue &&
                    element.ParentElementIndex.Value >= 0 &&
                    element.ParentElementIndex.Value < page.StructElements.Count
                        ? page.StructElements[element.ParentElementIndex.Value].ObjectId
                        : documentStructElementId;
                tableHeaders.TryGetValue(element, out var headers);
                string structElement;
                if (element.AnnotationObjectId.HasValue) {
                    structElement = PdfStructTreeRootDictionaryBuilder.BuildAnnotationStructElement(
                        parentObjectId,
                        pageIds[pageIndex],
                        element.AnnotationObjectId.Value,
                        element.MarkedContentId,
                        element.AdditionalMarkedContentIds,
                        element.AdditionalAnnotationObjectIds,
                        element.StructureType,
                        element.AlternativeText,
                        contentStreamObjectId,
                        additionalContentStreamObjectIds,
                        childElementIds: structureChildren.GetChildren(page, element, elementIndex));
                } else if (element.MarkedContentId.HasValue) {
                    structElement = string.Equals(element.StructureType, "Figure", StringComparison.Ordinal)
                        ? PdfStructTreeRootDictionaryBuilder.BuildFigureStructElement(
                            parentObjectId,
                            pageIds[pageIndex],
                            element.MarkedContentId.Value,
                            element.AlternativeText,
                            contentStreamObjectId)
                        : PdfStructTreeRootDictionaryBuilder.BuildTextStructElement(
                            parentObjectId,
                            pageIds[pageIndex],
                            element.StructureType,
                            element.MarkedContentId.Value,
                            element.TableHeaderScope,
                            element.TableColumnSpan,
                            element.TableRowSpan,
                            element.AdditionalMarkedContentIds,
                            contentStreamObjectId,
                            additionalContentStreamObjectIds, tableHeaders: headers);
                } else {
                    structElement = PdfStructTreeRootDictionaryBuilder.BuildContainerStructElement(
                        parentObjectId,
                        pageIds[pageIndex],
                        element.StructureType,
                        structureChildren.GetChildren(page, element, elementIndex),
                        element.TableHeaderScope,
                        element.TableColumnSpan,
                        element.TableRowSpan,
                        element.AlternativeText,
                        includePageReference: !element.SpansPages,
                        associatedFileIds: element.AssociatedFileIds,
                        tableHeaders: headers);
                }

                string? language = element.Language;
                if (string.Equals(language, InheritedStructureLanguage(element, structureParents, documentLanguage), StringComparison.OrdinalIgnoreCase)) language = null;
                structElement = PdfStructTreeRootDictionaryBuilder.WithStructureMetadata(structElement, namespaceId, language, StructureId(element));
                if (element.StructureType == "L" && element.ListNumbering.HasValue) {
                    PdfListNumbering numbering = element.ListNumbering.Value;
                    if (!pdf20 && numbering >= PdfListNumbering.Ordered)
                        throw new NotSupportedException("The selected list numbering attribute requires PDF 2.0.");
                    structElement = structElement.Insert(structElement.LastIndexOf(">>", StringComparison.Ordinal),
                        " /A << /O /List /ListNumbering /" + (numbering == PdfListNumbering.DecimalNumbers ? "Decimal" : numbering.ToString()) + " >> ");
                }
                ReplaceObject(objects, element.ObjectId, structElement);
            }

            var pageMarkedContentElements = new List<(int MarkedContentId, int ObjectId)>();
            foreach (PageStructElement element in page.StructElements.Where(element => element.MarkedContentId.HasValue)) {
                pageMarkedContentElements.Add((element.MarkedContentId!.Value, element.ObjectId));
                if (element.AdditionalMarkedContentIds != null) {
                    for (int additionalIndex = 0; additionalIndex < element.AdditionalMarkedContentIds.Count; additionalIndex++) {
                        pageMarkedContentElements.Add((element.AdditionalMarkedContentIds[additionalIndex], element.ObjectId));
                    }
                }
            }

            var pageOwnedMappings = pageMarkedContentElements
                .Where(mapping => FindMarkedContentStreamObjectId(page, mapping.MarkedContentId) == null)
                .ToList();
            var pageElementIds = BuildMarkedContentParentArray(pageOwnedMappings);

            for (int elementIndex = 0; elementIndex < page.StructElements.Count; elementIndex++) {
                PageStructElement element = page.StructElements[elementIndex];
                if (!element.ParentElementIndex.HasValue && element.ParentElement == null) {
                    documentChildren.Add(element);
                }
            }

            if (page.StructParentIndex.HasValue && pageElementIds.Count > 0) {
                parentTreeEntries.Add(PdfStructTreeRootDictionaryBuilder.ParentTreeEntry.ForMarkedContentPage(page.StructParentIndex.Value, pageElementIds));
            }

            foreach (PageEffectGroup effect in page.EffectGroups.Where(effect => effect.StructParentIndex.HasValue && effect.MarkedContentIds.Count > 0)) {
                var effectMappings = pageMarkedContentElements
                    .Where(mapping => effect.MarkedContentIds.Contains(mapping.MarkedContentId))
                    .ToList();
                var effectElementIds = BuildMarkedContentParentArray(effectMappings);
                if (effectElementIds.Count > 0) {
                    parentTreeEntries.Add(PdfStructTreeRootDictionaryBuilder.ParentTreeEntry.ForMarkedContentContainer(effect.StructParentIndex!.Value, effectElementIds));
                }
            }

            foreach (PageStructElement element in page.StructElements.Where(element => element.AnnotationObjectId.HasValue && element.AnnotationStructParentIndex.HasValue).OrderBy(element => element.AnnotationStructParentIndex!.Value)) {
                parentTreeEntries.Add(PdfStructTreeRootDictionaryBuilder.ParentTreeEntry.ForObjectReference(element.AnnotationStructParentIndex!.Value, element.ObjectId));
                if (element.AdditionalAnnotationStructParentIndexes != null) {
                    for (int additionalIndex = 0; additionalIndex < element.AdditionalAnnotationStructParentIndexes.Count; additionalIndex++) {
                        parentTreeEntries.Add(PdfStructTreeRootDictionaryBuilder.ParentTreeEntry.ForObjectReference(element.AdditionalAnnotationStructParentIndexes[additionalIndex], element.ObjectId));
                    }
                }
            }
        }

        List<int> documentChildElementIds = OrderStructureChildren(documentChildren);
        string documentElement = PdfStructTreeRootDictionaryBuilder.BuildDocumentStructElement(structTreeRootId, documentChildElementIds, documentLanguage);
        ReplaceObject(objects, documentStructElementId, PdfStructTreeRootDictionaryBuilder.WithStructureMetadata(documentElement, namespaceId));
        int parentTreeId = AddObject(objects, PdfStructTreeRootDictionaryBuilder.BuildParentTree(parentTreeEntries, value => AddObject(objects, value)));
        int parentTreeNextKey = parentTreeEntries.Count == 0
            ? 0
            : parentTreeEntries.Max(entry => entry.StructParentIndex) + 1;
        ReplaceObject(objects, structTreeRootId, PdfStructTreeRootDictionaryBuilder.BuildStructTreeRootDictionary(new[] { documentStructElementId }, parentTreeId, parentTreeNextKey,
            pages.SelectMany(page => page.StructElements).Select(element => element.StructureType), BuildStructureIdTree(objects, pages), namespaceId));
    }

    // A paragraph/cell/heading with interleaved links owns ordered Span/Link children.
    // Keeping all text MCIDs on its original leaf would move trailing text before links.
    private static void PromoteTextStructureContainer(LayoutResult.Page page, int? index) {
        if (!index.HasValue) return;
        PageStructElement element = page.StructElements[index.Value];
        if (!element.MarkedContentId.HasValue) return;
        page.StructElements.Add(new PageStructElement {
            MarkedContentId = element.MarkedContentId,
            AdditionalMarkedContentIds = element.AdditionalMarkedContentIds,
            StructureType = "Span",
            ParentElementIndex = index
        });
        element.MarkedContentId = null;
        element.AdditionalMarkedContentIds = null;
    }

    // Object numbers provide document-local uniqueness, including notes repeated across pages.
    private static string NoteStructureId(int objectId) => "note-" + objectId.ToString(System.Globalization.CultureInfo.InvariantCulture);

    private static int? BuildStructureIdTree(IList<byte[]> objects, IReadOnlyList<LayoutResult.Page> pages) {
        var notes = pages.SelectMany(page => page.StructElements)
            .Where(element => StructureId(element) != null)
            .OrderBy(element => StructureId(element), StringComparer.Ordinal).ToArray();
        if (notes.Length == 0) return null;
        var names = new StringBuilder("<< /Names [");
        foreach (var note in notes) {
            names.Append(PdfSyntaxEscaper.TextString(StructureId(note)!)).Append(' ')
                .Append(PdfSyntaxEscaper.IndirectReference(note.ObjectId)).Append(' ');
        }
        return AddObject(objects, names.Append("] >>\n").ToString());
    }
}
