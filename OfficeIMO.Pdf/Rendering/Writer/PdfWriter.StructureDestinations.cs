namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Completes logical link targets after generated structure object IDs are assigned.</summary>
    private static void CompleteStructureDestinations(IList<byte[]> objects, IReadOnlyList<LayoutResult.Page> pages, bool required) {
        var targets = pages.SelectMany(page => page.NamedDestinations)
            .ToDictionary(destination => destination.Name, StringComparer.Ordinal);
        foreach (var page in pages) {
            foreach (var link in page.Annotations) {
                if (string.IsNullOrEmpty(link.DestinationName)) continue;
                PageStructElement? target = targets[link.DestinationName!].StructureElement;
                int? structureId = target?.ObjectId > 0 ? target.ObjectId : null;
                if (!structureId.HasValue && required)
                    throw new InvalidOperationException("PDF/UA-2 internal links require a named destination inside tagged content: " + link.DestinationName);
                ReplaceObject(objects, link.ObjectId, PdfAnnotationDictionaryBuilder.BuildGoToNamedDestinationLinkAnnotation(
                    link.X1, link.Y1, link.X2, link.Y2, link.DestinationName!, link.Contents, link.StructParentIndex, structureId));
            }
        }
    }

    private static int? ResolveOutlineStructureId(LayoutResult.Page page, OutlineNode outline, bool required) {
        PageStructElement? target = outline.StructureElement;
        if (target == null && outline.StructureElementKey != null)
            target = page.StructElements.FirstOrDefault(element => element.StructureElementKey == outline.StructureElementKey);
        if (target?.ObjectId > 0) return target.ObjectId;
        if (required) throw new InvalidOperationException("PDF/UA-2 outlines require a tagged target. Add canvas outlines inside their Structure container.");
        return null;
    }

    private sealed partial class LayoutContext {
        // Deferred figures/widgets still reserve their source identity before later text is laid out.
        // The same object is added to the structure tree when its paint/annotation is serialized.
        private PageStructElement? ReservePendingStructureDestination(string role) {
            if (_suppressCanvasStructureRegistration || !emitGeneratedStructure || currentPage == null ||
                !currentPage.NamedDestinations.Any(destination => destination.AttachToNextStructure)) return null;
            var target = new PageStructElement { StructureType = role };
            BindPendingStructureDestinations(target);
            return target;
        }

        private void ReserveImageNavigationTarget(PageImage image) {
            if (!image.SuppressAccessibilityWrapper && !image.IsBackgroundDecoration && !string.IsNullOrWhiteSpace(image.AlternativeText))
                image.NavigationTarget = ReservePendingStructureDestination("Figure");
        }

        private void CaptureFormNavigationTarget(FormFieldAnnotation field) {
            field.NavigationTarget = ReservePendingStructureDestination("Form");
            if (field.RadioWidgets.Count > 0) field.RadioWidgets[0].NavigationTarget = field.NavigationTarget;
        }

        private void AddFormFieldAnnotation(FormFieldAnnotation field) {
            RecordFlowPlacement(field.Y2);
            CaptureFormNavigationTarget(field);
            currentPage!.FormFields.Add(field);
        }

        /// <summary>Retargets a clipped flow destination to its first surviving source admission, including siblings already laid out.</summary>
        private void ReleasePendingStructureDestination(PageStructElement? element) {
            if (element == null || currentPage == null) return;
            element.NavigationSourceRemoved = true;
            PageStructElement? next = element.NextNavigationSource;
            while (next?.NavigationSourceRemoved == true) next = next.NextNavigationSource;
            // Collapse discarded chains before the next reverse clipping removal.
            element.NextNavigationSource = next;
            foreach (var destination in currentPage.NamedDestinations) {
                if (destination.AttachToNextStructure && ReferenceEquals(destination.StructureElement, element))
                    destination.StructureElement = next;
            }
        }

        private void BindPendingStructureDestinations(PageStructElement element) {
            if (currentPage == null) return;
            if (currentPage.LastNavigationSource is { } previous && !ReferenceEquals(previous, element))
                previous.NextNavigationSource = element;
            currentPage.LastNavigationSource = element;
            foreach (var destination in currentPage.NamedDestinations) {
                if (destination.AttachToNextStructure && destination.StructureElement == null)
                    destination.StructureElement = element;
            }
        }

        private PageStructElement? ResolveMarkedStructure(int? markedContentId) {
            if (!markedContentId.HasValue || currentPage == null) return null;
            PageStructElement? element = currentPage.StructElements.FirstOrDefault(item => item.MarkedContentId == markedContentId);
            if (element == null) return null;
            PageStructElement? parent = element.ParentElement;
            if (parent == null && element.ParentElementIndex is int index) parent = currentPage.StructElements[index];
            return parent != null && parent.StructureType.Length > 0 && parent.StructureType[0] == 'H' ? parent : element;
        }


    }
}
