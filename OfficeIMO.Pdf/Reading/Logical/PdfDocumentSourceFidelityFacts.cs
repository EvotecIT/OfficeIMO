namespace OfficeIMO.Pdf;

/// <summary>
/// Counts document-wide source features for loss reporting after page selection has
/// intentionally removed their details from the logical projection.
/// </summary>
internal readonly struct PdfDocumentSourceFidelityFacts {
    private readonly FormFieldScope[] _formFields;

    internal PdfDocumentSourceFidelityFacts(PdfReadDocument document, int[] selectedPageNumbers) {
        _formFields = document.FormFields.Select(static field => new FormFieldScope(field)).ToArray();
        AttachmentCount = document.Attachments.Count;
        CatalogActionCount = document.CatalogActions.Count;
        HasOpenAction = document.OpenAction != null;
        CatalogContainsOpenAction = document.CatalogActions.Any(static action =>
            string.Equals(action.Source, "OpenAction", StringComparison.Ordinal) && !action.IsChainedAction);
        UnplacedFormFieldCount = document.FormFields.Count(static field => field.HasUnplacedContent);
        FieldWithoutWidgetCount = document.FormFields.Count(static field => field.Widgets.Count == 0);
        RelevantFormFieldCount = CountRelevantFormFields(_formFields, selectedPageNumbers, document.Pages.Count);
        OptionalContentGroupCount = document.OptionalContent?.Groups.Count ?? 0;
        HasTaggedContent = document.TaggedContent != null;
        HasDocumentMetadata = document.UncheckedMetadata.HasContent || document.UncheckedXmpMetadata != null;
        StructureElementCount = document.TaggedContent?.StructureElementCount ?? 0;
        MarkedContentReferenceCount = document.TaggedContent?.MarkedContentReferenceCount ?? 0;
    }

    private PdfDocumentSourceFidelityFacts(PdfDocumentSourceFidelityFacts source, int relevantFormFieldCount) {
        _formFields = source._formFields;
        AttachmentCount = source.AttachmentCount;
        CatalogActionCount = source.CatalogActionCount;
        HasOpenAction = source.HasOpenAction;
        CatalogContainsOpenAction = source.CatalogContainsOpenAction;
        UnplacedFormFieldCount = source.UnplacedFormFieldCount;
        FieldWithoutWidgetCount = source.FieldWithoutWidgetCount;
        RelevantFormFieldCount = relevantFormFieldCount;
        OptionalContentGroupCount = source.OptionalContentGroupCount;
        HasTaggedContent = source.HasTaggedContent;
        HasDocumentMetadata = source.HasDocumentMetadata;
        StructureElementCount = source.StructureElementCount;
        MarkedContentReferenceCount = source.MarkedContentReferenceCount;
    }

    internal PdfDocumentSourceFidelityFacts ForPageNumbers(int[] selectedPageNumbers, int sourcePageCount) =>
        new(this, CountRelevantFormFields(_formFields, selectedPageNumbers, sourcePageCount));

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
    internal bool HasDocumentMetadata { get; }
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
