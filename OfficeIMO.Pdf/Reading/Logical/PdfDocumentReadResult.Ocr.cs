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
        return new PdfDocumentReadResult(
            Metadata,
            pages,
            Outlines,
            PageLabels,
            NamedDestinations,
            CatalogActions,
            Attachments,
            OutputIntents,
            XmpMetadata,
            TaggedContent,
            OptionalContent,
            OpenAction,
            ViewerPreferences,
            FormFields,
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
            SourcePageCount,
            Profile);
    }
}
