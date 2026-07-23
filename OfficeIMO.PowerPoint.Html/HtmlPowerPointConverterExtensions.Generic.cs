using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class HtmlPowerPointConverterExtensions {
    private static void ImportGenericDocument(
        HtmlSemanticDocument document,
        PptCore.PowerPointPresentation presentation,
        HtmlToPowerPointOptions options,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget) {
        foreach (HtmlSemanticSection section in document.Sections) {
            if (!budget.TryReserveSemanticContainer(out string containerLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "Additional HTML sections were omitted because the shared slide limit was reached.",
                    HtmlDiagnosticSeverity.Error, HtmlConversionLossKind.Omission, detail: containerLimit);
                break;
            }

            PptCore.PowerPointSlide slide = presentation.AddSlide();
            result.Slides++;
            double contentTop = 30D;
            if (!string.IsNullOrWhiteSpace(section.Title)) {
                HtmlSemanticBlock? titleBlock = section.Blocks.FirstOrDefault();
                if (titleBlock != null) {
                    contentTop = ImportTextBox(titleBlock.SourceElement, section.Title, slide, 30D, result, budget, 44D);
                }
            }

            double pictureTop = contentTop;
            foreach (HtmlSemanticBlock block in section.Blocks) {
                bool isSectionTitle = block.Kind == HtmlSemanticBlockKind.Heading
                    && string.Equals(block.Text, section.Title, StringComparison.Ordinal);
                bool importText = IsGenericTextBlock(block.Kind);
                bool importTable = options.ImportTables && block.Kind == HtmlSemanticBlockKind.Table;
                bool importPicture = options.ImportPictures && block.Kind == HtmlSemanticBlockKind.Image;
                if (importText && !isSectionTitle) {
                    contentTop = ImportTextBox(block.SourceElement, block.Text, slide, contentTop, result, budget,
                        block.Kind == HtmlSemanticBlockKind.List ? Math.Max(52D, block.Children.Count * 30D) : 52D,
                        block);
                } else if (importTable) {
                    contentTop = ImportTable(block.SourceElement, slide, contentTop, result, budget, block);
                } else if (importPicture) {
                    pictureTop = Math.Max(pictureTop, contentTop);
                    ImportPicture(block.SourceElement, slide, result, budget, ref pictureTop);
                    contentTop = Math.Max(contentTop, pictureTop);
                }
                if (options.ImportPictures) {
                    foreach (HtmlSemanticResource resource in EnumerateInlineResources(block)) {
                        pictureTop = Math.Max(pictureTop, contentTop);
                        ImportSemanticResourcePicture(resource, slide, result, budget, ref pictureTop);
                        contentTop = Math.Max(contentTop, pictureTop);
                    }
                }
            }
        }
    }

    private static IEnumerable<HtmlSemanticResource> EnumerateInlineResources(HtmlSemanticBlock block) {
        foreach (HtmlSemanticResource resource in block.InlineResources.Where(item => item.Kind == HtmlResourceKind.Image)) yield return resource;
        if (block.Table != null) {
            foreach (HtmlSemanticResource resource in block.Table.Rows.SelectMany(row => row.Cells)
                .SelectMany(cell => cell.Resources).Where(item => item.Kind == HtmlResourceKind.Image)) yield return resource;
        }
        foreach (HtmlSemanticBlock child in block.Children) {
            foreach (HtmlSemanticResource resource in EnumerateInlineResources(child)) yield return resource;
        }
    }

    private static void ImportSemanticResourcePicture(
        HtmlSemanticResource resource,
        PptCore.PowerPointSlide slide,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget,
        ref double top) {
        if (!HtmlImageDataUri.TryParse(resource.Source, out HtmlImageDataUri dataUri)
            || !TryGetImagePartType(dataUri.MediaType, out PptCore.ImagePartType imagePartType)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceTypeUnsupported,
                "An inline generic slide image was omitted because native import requires a supported bounded image data URI.",
                lossKind: HtmlConversionLossKind.Omission, source: resource.Source);
            return;
        }
        if (!budget.IsImageWithinLimit(dataUri, out string limit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "An inline generic slide image was omitted because it exceeded the shared image limit.",
                lossKind: HtmlConversionLossKind.Omission, source: resource.Source, detail: limit);
            return;
        }
        if (!budget.TryReserveImageWithShape(dataUri, out limit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "An inline generic slide image was omitted because the shared image or shape limit was reached.",
                lossKind: HtmlConversionLossKind.Omission, source: resource.Source, detail: limit);
            return;
        }
        if (!dataUri.TryDecodeBytes(out byte[] bytes)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceDecodeFailed,
                "An inline generic slide image could not be decoded.",
                lossKind: HtmlConversionLossKind.Omission, source: resource.Source);
            return;
        }
        double maximum = budget.Limits.MaxAbsoluteGeometry;
        double width = Math.Min(maximum, Math.Max(1D, (resource.WidthPixels ?? 160D) * 0.75D));
        double height = Math.Min(maximum, Math.Max(1D, (resource.HeightPixels ?? 90D) * 0.75D));
        using var stream = new MemoryStream(bytes);
        PptCore.PowerPointPicture picture = slide.AddPicturePoints(stream, imagePartType, 64D, top, width, height);
        if (!string.IsNullOrWhiteSpace(resource.AlternateText)) picture.AltText = resource.AlternateText;
        result.Pictures++;
        top += height + 18D;
    }

    private static bool IsGenericTextBlock(HtmlSemanticBlockKind kind) =>
        kind == HtmlSemanticBlockKind.Heading || kind == HtmlSemanticBlockKind.Paragraph
        || kind == HtmlSemanticBlockKind.Code || kind == HtmlSemanticBlockKind.Quote
        || kind == HtmlSemanticBlockKind.List || kind == HtmlSemanticBlockKind.Note;
}
