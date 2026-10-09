using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public static partial class OfficeDocumentModelPdfExtensions {
    private static bool HasExplicitReadingOrder(OfficeDocumentModel source) =>
        source.Blocks.Any(block => block.Location?.LogicalOrder.HasValue == true)
        || source.Tables.Any(table => table.Location?.LogicalOrder.HasValue == true)
        || source.Pages.Any(page => page.Blocks.Any(block => block.Location?.LogicalOrder.HasValue == true)
            || page.Tables.Any(table => table.Location?.LogicalOrder.HasValue == true));

    private static AssetProjectionSummary ComposeReadingOrderedContent(PdfDocument document, OfficeDocumentModel source,
        ProjectionIdentitySet identities, PdfProjectionOptions options, OfficeRasterDecodeOptions rasterDecodeOptions,
        bool allowRelativeUriLinks, PdfConversionReport report, System.Threading.CancellationToken token) {
        var blocks = new List<OfficeDocumentModelBlock>();
        var tables = new List<OfficeDocumentModelTable>();
        var assets = new List<OfficeDocumentModelAsset>();
        var links = new List<OfficeDocumentModelLink>();
        var forms = new List<OfficeDocumentModelFormField>();
        bool omittedPageNames = false;
        foreach (OfficeDocumentModelPage page in source.Pages) {
            token.ThrowIfCancellationRequested();
            blocks.AddRange(identities.TakeBlocks(page.Blocks));
            tables.AddRange(identities.TakeTables(page.Tables, page));
            assets.AddRange(identities.TakeAssets(page.Assets));
            links.AddRange(identities.TakeLinks(page.Links));
            forms.AddRange(identities.TakeForms(page.Forms));
            omittedPageNames |= !string.IsNullOrWhiteSpace(page.Name);
        }
        blocks.AddRange(identities.TakeBlocks(source.Blocks));
        tables.AddRange(identities.TakeTables(source.Tables));
        assets.AddRange(identities.TakeAssets(source.Assets));
        links.AddRange(identities.TakeLinks(source.Links));
        forms.AddRange(identities.TakeForms(source.Forms));
        if (omittedPageNames) {
            report.Add(new PdfConversionWarning(ConverterName, "MODEL_SOURCE_PAGE_LABELS_OMITTED", source.Format + "/document",
                "Automatic physical-page headings are omitted in continuous flow to preserve the source's explicit logical reading order.",
                PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission));
        }
        return ComposeContent(document, blocks, tables, assets, links, forms, options, rasterDecodeOptions,
            allowRelativeUriLinks, report, source.Format + "/document", token);
    }
}
