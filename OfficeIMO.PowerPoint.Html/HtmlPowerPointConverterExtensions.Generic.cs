using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class HtmlPowerPointConverterExtensions {
    private static void ImportGenericDocument(
        IHtmlDocument document,
        PptCore.PowerPointPresentation presentation,
        HtmlToPowerPointOptions options,
        HtmlToPowerPointResult result,
        HtmlImportBudget budget) {
        foreach (HtmlGenericSectionProjection section in HtmlGenericDocumentProjector.CreateSections(document)) {
            if (!budget.TryReserveSemanticContainer(out string containerLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "Additional HTML sections were omitted because the shared slide limit was reached.",
                    HtmlDiagnosticSeverity.Error, HtmlConversionLossKind.Omission, detail: containerLimit);
                break;
            }

            PptCore.PowerPointSlide slide = presentation.AddSlide();
            result.Slides++;
            double contentTop = 30D;
            if (!string.IsNullOrWhiteSpace(section.Title) && budget.TryReserveShape(out _)) {
                contentTop = ImportTextBox(section.Blocks.FirstOrDefault() ?? document.Body!, section.Title, slide, 30D, result, budget, 44D);
            }

            double pictureTop = contentTop;
            foreach (IElement block in HtmlGenericDocumentProjector.EnumerateBlocks(section)) {
                if (HtmlGenericDocumentProjector.IsHeading(block)
                    && string.Equals(HtmlGenericDocumentProjector.GetBlockText(block), section.Title, StringComparison.Ordinal)) {
                    continue;
                }
                bool importText = HtmlGenericDocumentProjector.IsTextBlock(block);
                bool importTable = options.ImportTables && HtmlGenericDocumentProjector.IsTable(block);
                bool importPicture = options.ImportPictures && HtmlGenericDocumentProjector.IsImage(block);
                if (!importText && !importTable && !importPicture) continue;
                if (!budget.TryReserveShape(out string shapeLimit)) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "Additional HTML blocks were omitted because the shared shape limit was reached.",
                        lossKind: HtmlConversionLossKind.Omission, detail: shapeLimit);
                    break;
                }

                if (importText) {
                    string text = HtmlGenericDocumentProjector.GetBlockText(block);
                    contentTop = ImportTextBox(block, text, slide, contentTop, result, budget, 52D);
                } else if (importTable) {
                    contentTop = ImportTable(block, slide, contentTop, result, budget);
                } else if (importPicture) {
                    pictureTop = Math.Max(pictureTop, contentTop);
                    ImportPicture(block, slide, result, budget, ref pictureTop);
                    contentTop = Math.Max(contentTop, pictureTop);
                }
            }
        }
    }
}
