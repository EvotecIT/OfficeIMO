namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentReadResult {
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
            Profile);
    }
}
