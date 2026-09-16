namespace OfficeIMO.Pdf;

/// <summary>
/// Counts document-wide source features for loss reporting after page selection has
/// intentionally removed their details from the logical projection.
/// </summary>
internal readonly struct PdfDocumentSourceFidelityFacts {
    internal PdfDocumentSourceFidelityFacts(PdfReadDocument document, int[] selectedPageNumbers) {
        AttachmentCount = document.Attachments.Count;
        CatalogActionCount = document.CatalogActions.Count;
        HasOpenAction = document.OpenAction != null;
        CatalogContainsOpenAction = document.CatalogActions.Any(static action =>
            string.Equals(action.Source, "OpenAction", StringComparison.Ordinal) && !action.IsChainedAction);
        UnplacedFormFieldCount = document.FormFields.Count(static field => field.HasUnplacedContent);
        FieldWithoutWidgetCount = document.FormFields.Count(static field => field.Widgets.Count == 0);
        if (PdfPageRangeObjectFilter.ShouldUseDocumentWideObjects(document.Pages.Count, selectedPageNumbers)) {
            RelevantFormFieldCount = document.FormFields.Count;
        } else {
            var selectedPages = new HashSet<int>(selectedPageNumbers);
            RelevantFormFieldCount = document.FormFields.Count(field =>
                field.HasUnplacedContent || field.Widgets.Any(widget =>
                    widget.PageNumber.HasValue && selectedPages.Contains(widget.PageNumber.Value)));
        }
        OptionalContentGroupCount = document.OptionalContent?.Groups.Count ?? 0;
        HasTaggedContent = document.TaggedContent != null;
        StructureElementCount = document.TaggedContent?.StructureElementCount ?? 0;
        MarkedContentReferenceCount = document.TaggedContent?.MarkedContentReferenceCount ?? 0;
    }

    internal int AttachmentCount { get; }
    internal int CatalogActionCount { get; }
    internal bool HasOpenAction { get; }
    internal bool CatalogContainsOpenAction { get; }
    internal int UnplacedFormFieldCount { get; }
    internal int FieldWithoutWidgetCount { get; }
    internal int RelevantFormFieldCount { get; }
    internal int OptionalContentGroupCount { get; }
    internal bool HasTaggedContent { get; }
    internal int StructureElementCount { get; }
    internal int MarkedContentReferenceCount { get; }
}
