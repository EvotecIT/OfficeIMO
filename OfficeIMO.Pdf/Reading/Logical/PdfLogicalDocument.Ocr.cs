namespace OfficeIMO.Pdf;

public sealed partial class PdfLogicalDocument {
    internal PdfLogicalDocument WithPages(IReadOnlyList<PdfLogicalPage> pages) {
        return new PdfLogicalDocument(
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
            CatalogLanguage);
    }
}
