namespace OfficeIMO.Pdf;

/// <summary>
/// Counts document-wide source features for loss reporting after page selection has
/// intentionally removed their details from the logical projection.
/// </summary>
internal readonly struct PdfDocumentSourceFidelityFacts {
    private readonly FormFieldScope[] _formFields;
    private readonly PdfDocumentOpenAction? _openAction;
    private readonly int _openActionCatalogActionCount;

    internal PdfDocumentSourceFidelityFacts(PdfReadDocument document, int[] selectedPageNumbers) {
        _formFields = document.FormFields.Select(static field => new FormFieldScope(field)).ToArray();
        _openAction = PdfPageRangeObjectFilter.ShouldUseDocumentWideObjects(document.Pages.Count, selectedPageNumbers)
            ? document.OpenAction
            : PdfPageRangeObjectFilter.FilterOpenActionByPageNumbers(document.OpenAction, selectedPageNumbers);
        _openActionCatalogActionCount = document.CatalogActions.Count(static action =>
            string.Equals(action.Source, "OpenAction", StringComparison.Ordinal));
        bool excludedReadableOpenAction = document.OpenAction != null && _openAction == null;
        AttachmentCount = document.Attachments.Count;
        CatalogActionCount = Math.Max(
            0,
            document.CatalogActions.Count - (excludedReadableOpenAction ? _openActionCatalogActionCount : 0));
        HasOpenAction = _openAction != null;
        CatalogContainsOpenAction = !excludedReadableOpenAction && document.CatalogActions.Any(static action =>
            string.Equals(action.Source, "OpenAction", StringComparison.Ordinal) && !action.IsChainedAction);
        UnplacedFormFieldCount = document.FormFields.Count(static field => field.HasUnplacedContent);
        FieldWithoutWidgetCount = document.FormFields.Count(static field => field.Widgets.Count == 0);
        RelevantFormFieldCount = CountRelevantFormFields(_formFields, selectedPageNumbers, document.Pages.Count);
        OptionalContentGroupCount = document.OptionalContent?.Groups.Count ?? 0;
        HasTaggedContent = document.TaggedContent != null;
        HasInfoMetadata = document.UncheckedMetadata.HasContent;
        HasXmpMetadata = document.UncheckedXmpMetadata != null;
        StructureElementCount = document.TaggedContent?.StructureElementCount ?? 0;
        MarkedContentReferenceCount = document.TaggedContent?.MarkedContentReferenceCount ?? 0;
    }

    private PdfDocumentSourceFidelityFacts(
        PdfDocumentSourceFidelityFacts source,
        int relevantFormFieldCount,
        PdfDocumentOpenAction? openAction) {
        _formFields = source._formFields;
        _openAction = openAction;
        _openActionCatalogActionCount = source._openActionCatalogActionCount;
        AttachmentCount = source.AttachmentCount;
        bool excludedReadableOpenAction = source._openAction != null && openAction == null;
        CatalogActionCount = excludedReadableOpenAction
            ? Math.Max(0, source.CatalogActionCount - source._openActionCatalogActionCount)
            : source.CatalogActionCount;
        HasOpenAction = openAction != null;
        CatalogContainsOpenAction = !excludedReadableOpenAction && source.CatalogContainsOpenAction;
        UnplacedFormFieldCount = source.UnplacedFormFieldCount;
        FieldWithoutWidgetCount = source.FieldWithoutWidgetCount;
        RelevantFormFieldCount = relevantFormFieldCount;
        OptionalContentGroupCount = source.OptionalContentGroupCount;
        HasTaggedContent = source.HasTaggedContent;
        HasInfoMetadata = source.HasInfoMetadata;
        HasXmpMetadata = source.HasXmpMetadata;
        StructureElementCount = source.StructureElementCount;
        MarkedContentReferenceCount = source.MarkedContentReferenceCount;
    }

    internal PdfDocumentSourceFidelityFacts ForPageNumbers(int[] selectedPageNumbers, int sourcePageCount) =>
        new(
            this,
            CountRelevantFormFields(_formFields, selectedPageNumbers, sourcePageCount),
            PdfPageRangeObjectFilter.ShouldUseDocumentWideObjects(sourcePageCount, selectedPageNumbers)
                ? _openAction
                : PdfPageRangeObjectFilter.FilterOpenActionByPageNumbers(_openAction, selectedPageNumbers));

    private static int CountRelevantFormFields(FormFieldScope[] fields, int[] selectedPageNumbers, int sourcePageCount) {
        if (PdfPageRangeObjectFilter.ShouldUseDocumentWideObjects(sourcePageCount, selectedPageNumbers)) return fields.Length;
        var selectedPages = new HashSet<int>(selectedPageNumbers);
        return fields.Count(field => field.HasUnplacedContent || field.PageNumbers.Any(selectedPages.Contains));
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
    internal bool HasInfoMetadata { get; }
    internal bool HasXmpMetadata { get; }
    internal bool HasDocumentMetadata => HasInfoMetadata || HasXmpMetadata;
    internal int StructureElementCount { get; }
    internal int MarkedContentReferenceCount { get; }

    private readonly struct FormFieldScope {
        internal FormFieldScope(PdfFormField field) {
            HasUnplacedContent = field.HasUnplacedContent;
            PageNumbers = HasUnplacedContent
                ? Array.Empty<int>()
                : field.Widgets.Where(static widget => widget.PageNumber.HasValue)
                    .Select(static widget => widget.PageNumber!.Value)
                    .Distinct()
                    .ToArray();
        }

        internal bool HasUnplacedContent { get; }
        internal int[] PageNumbers { get; }
    }
}
