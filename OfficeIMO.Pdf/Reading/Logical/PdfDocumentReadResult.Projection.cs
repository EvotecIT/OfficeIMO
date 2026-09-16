namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentReadResult {
    internal PdfDocumentReadResult ProjectPages(
        PdfPageSelection? selection,
        string parameterName,
        System.Threading.CancellationToken cancellationToken = default) {
        if (selection is null) return this;

        int[] pageNumbers = selection.ToPageNumbers(SourcePageCount, parameterName);
        var pages = new List<PdfLogicalPage>(pageNumbers.Length);
        for (int pageIndex = 0; pageIndex < pageNumbers.Length; pageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            int pageNumber = pageNumbers[pageIndex];
            IReadOnlyList<PdfLogicalPage> occurrences = GetPages(pageNumber);
            if (occurrences.Count == 0) {
                throw new ArgumentOutOfRangeException(
                    parameterName,
                    pageNumber,
                    "The PDF page selection references a source page that is not present in this reconstructed result.");
            }

            for (int occurrenceIndex = 0; occurrenceIndex < occurrences.Count; occurrenceIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                pages.Add(occurrences[occurrenceIndex]);
            }
        }

        return WithPages(pages.AsReadOnly());
    }

    internal PdfDocumentReadResult WithPages(IReadOnlyList<PdfLogicalPage> pages) {
        int[] selectedPageNumbers = pages.Select(static page => page.PageNumber).ToArray();
        PdfDocumentSourceFidelityFacts sourceFidelityFacts = SourceFidelityFacts.ForPageNumbers(
            selectedPageNumbers,
            SourcePageCount);
        bool useDocumentWideObjects = PdfPageRangeObjectFilter.ShouldUseDocumentWideObjects(
            SourcePageCount, selectedPageNumbers);
        return new PdfDocumentReadResult(
            Metadata,
            pages,
            useDocumentWideObjects ? Outlines : PdfPageRangeObjectFilter.FilterOutlinesByPageNumbers(Outlines, selectedPageNumbers),
            useDocumentWideObjects ? PageLabels : PdfPageRangeObjectFilter.FilterPageLabelsByPageNumbers(PageLabels, selectedPageNumbers),
            useDocumentWideObjects ? NamedDestinations : PdfPageRangeObjectFilter.FilterNamedDestinationsByPageNumbers(NamedDestinations, selectedPageNumbers),
            useDocumentWideObjects ? CatalogActions : Array.Empty<PdfCatalogAction>(),
            useDocumentWideObjects ? Attachments : Array.Empty<PdfAttachmentInfo>(),
            useDocumentWideObjects ? OutputIntents : Array.Empty<PdfOutputIntentInfo>(),
            useDocumentWideObjects ? XmpMetadata : null,
            useDocumentWideObjects ? TaggedContent : null,
            useDocumentWideObjects ? OptionalContent : null,
            useDocumentWideObjects ? OpenAction : PdfPageRangeObjectFilter.FilterOpenActionByPageNumbers(OpenAction, selectedPageNumbers),
            ViewerPreferences,
            useDocumentWideObjects ? FormFields : PdfPageRangeObjectFilter.FilterFormFieldsByPageNumbers(FormFields, selectedPageNumbers, preservePageDuplicates: false),
            AcroFormDefaultAppearance,
            AcroFormQuadding,
            AcroFormXfa,
            AcroFormNeedAppearances,
            AcroFormSignatureFlags,
            Security,
            CatalogPageMode,
            CatalogPageLayout,
            CatalogVersion,
            CatalogLanguage,
            sourceFidelityFacts,
            SourcePageCount,
            Profile);
    }
}
