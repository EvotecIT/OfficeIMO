using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentReadResult {
    internal static PdfDocumentReadResult FromPageNumbers(
        PdfReadDocument document,
        PdfTextLayoutOptions? options,
        int[] pageNumbers,
        IReadOnlyList<PdfUnderstandingPageResult> analyses,
        PdfReadProfile profile,
        CancellationToken cancellationToken,
        bool preserveDocumentWideObjects = false, bool includeImages = true, bool includeFormWidgets = true) {
        document.DemandContentExtraction("logical object");
        if (analyses.Count != pageNumbers.Length) {
            throw new ArgumentException("Semantic analysis count must match the selected page count.", nameof(analyses));
        }
        bool useDocumentWideObjects = preserveDocumentWideObjects || PdfPageRangeObjectFilter.ShouldUseDocumentWideObjects(document.Pages.Count, pageNumbers);
        IReadOnlyList<PdfFormField> formFields = useDocumentWideObjects
            ? document.FormFields
            : PdfPageRangeObjectFilter.FilterFormFieldsByPageNumbers(document.FormFields, pageNumbers, preservePageDuplicates: false);
        IReadOnlyList<PdfOutlineItem> outlines = useDocumentWideObjects
            ? document.Outlines
            : PdfPageRangeObjectFilter.FilterOutlinesByPageNumbers(document.Outlines, pageNumbers);
        IReadOnlyList<PdfPageLabel> pageLabels = useDocumentWideObjects
            ? document.PageLabels
            : PdfPageRangeObjectFilter.FilterPageLabelsByPageNumbers(document.PageLabels, pageNumbers);
        IReadOnlyList<PdfNamedDestination> namedDestinations = useDocumentWideObjects
            ? document.NamedDestinations
            : PdfPageRangeObjectFilter.FilterNamedDestinationsByPageNumbers(document.NamedDestinations, pageNumbers);
        IReadOnlyList<PdfCatalogAction> catalogActions = useDocumentWideObjects
            ? document.CatalogActions
            : Array.Empty<PdfCatalogAction>();
        IReadOnlyList<PdfAttachmentInfo> attachments = useDocumentWideObjects
            ? document.Attachments
            : Array.Empty<PdfAttachmentInfo>();
        IReadOnlyList<PdfOutputIntentInfo> outputIntents = useDocumentWideObjects
            ? document.OutputIntents
            : Array.Empty<PdfOutputIntentInfo>();
        PdfXmpMetadataInfo? xmpMetadata = useDocumentWideObjects
            ? document.XmpMetadata
            : null;
        PdfTaggedContentInfo? taggedContent = useDocumentWideObjects
            ? document.TaggedContent
            : null;
        PdfOptionalContentProperties? optionalContent = useDocumentWideObjects
            ? document.OptionalContent
            : null;
        PdfDocumentOpenAction? openAction = useDocumentWideObjects
            ? document.OpenAction
            : PdfPageRangeObjectFilter.FilterOpenActionByPageNumbers(document.OpenAction, pageNumbers);
        IReadOnlyDictionary<int, IReadOnlyList<PdfLogicalFormWidget>> formWidgetsByPageNumber =
            includeFormWidgets ? IndexFormWidgetsByPageNumber(formFields) : new Dictionary<int, IReadOnlyList<PdfLogicalFormWidget>>();

        var pages = new List<PdfLogicalPage>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            int pageNumber = pageNumbers[i];
            IReadOnlyList<PdfLogicalFormWidget>? pageFormWidgets = null;
            if (formWidgetsByPageNumber.TryGetValue(pageNumber, out IReadOnlyList<PdfLogicalFormWidget>? sourceWidgets)) {
                pageFormWidgets = CloneFormWidgetsForOccurrence(sourceWidgets, pageNumber);
            }
            pages.Add(PdfLogicalPage.From(
                document,
                document.Pages[pageNumber - 1],
                pageNumber,
                options,
                pageFormWidgets,
                analyses[i],
                includeImages, includeFormWidgets, cancellationToken));
        }
        if (profile == PdfReadProfile.Structured) {
            ApplyDocumentHeadingFontTiers(
                pages.SelectMany(static page => page.Headings).ToArray(),
                cancellationToken);
        }

        cancellationToken.ThrowIfCancellationRequested();

        return new PdfDocumentReadResult(
            document.Metadata,
            pages.AsReadOnly(),
            outlines,
            pageLabels,
            namedDestinations,
            catalogActions,
            attachments,
            outputIntents,
            xmpMetadata,
            taggedContent,
            optionalContent,
            openAction,
            document.ViewerPreferences,
            formFields,
            document.AcroFormDefaultAppearance,
            document.AcroFormQuadding,
            document.AcroFormXfa,
            document.AcroFormNeedAppearances,
            document.AcroFormSignatureFlags,
            document.Security,
            document.CatalogPageMode,
            document.CatalogPageLayout,
            document.CatalogVersion,
            document.CatalogLanguage,
            new PdfDocumentSourceFidelityFacts(document, pageNumbers),
            document.Pages.Count,
            profile, includeImages, includeFormWidgets);
    }

}
