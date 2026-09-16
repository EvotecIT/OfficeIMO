using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

internal static partial class PdfWordConverter {
    internal static PdfWordConversionResult ConvertVisualPages(PdfCore.PdfDocument source, PdfToWordOptions options) {
        var token = options.CancellationToken;
        token.ThrowIfCancellationRequested();
        var renderOptions = new PdfCore.PdfPageRenderOptions {
            Dpi = options.Dpi, MaxPages = options.MaxPages, MaxPixelsPerPage = options.MaxPixelsPerPage,
            MaxOutputBytesPerPage = options.MaxOutputBytesPerPage, MaxTotalOutputBytes = options.MaxTotalOutputBytes,
            ContinueOnError = false
        };
        var pages = source.Render.Pages(options.ReadOptions?.PageSelection, renderOptions, token);
        if (pages.Count == 0) throw new InvalidOperationException("Select at least one PDF page for visual Word conversion.");
        WordDocument target = WordDocument.Create();
        try {
            PdfCore.PdfDocumentInfo sourceInfo = source.Inspect(null, token);
            PdfCore.PdfMetadata sourceMetadata = source.Reader.Metadata();
            if (options.IncludeMetadata) CopyMetadata(sourceMetadata, target);
            ReportMetadataFidelity(sourceMetadata, sourceInfo.HasXmpMetadata, options);
            ReportAttachmentsNotReconstructed(sourceInfo.AttachmentCount, options);
            if (sourceInfo.HasSecurityState) {
                AddWarning(options, "PdfSourceSecurityNotReconstructed", "Document/Security",
                    "PDF encryption, signature, permission, or revision state is not carried into the visual Word document.",
                    PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission);
            }
            if (sourceInfo.PageLabels.Any(label => pages.Any(page => page.PageNumber >= label.StartPageNumber))) {
                AddWarning(options, "PdfPageLabelsNotReconstructed", "Document/PageLabels",
                    "PDF page-label rules are not carried into the visual Word document.",
                    PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission);
            }
            if (sourceInfo.HasTaggedContent) {
                AddWarning(options, "PdfTaggedStructureNotReconstructed", "Document/StructTreeRoot",
                    "PDF tagged accessibility structure is not copied into the visual Word document.",
                    PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission);
            }
            if (sourceInfo.HasOpenActions || sourceInfo.CatalogActionCount > 0) {
                AddWarning(options, "PdfCatalogActionsNotReconstructed", "Document/CatalogActions",
                    "PDF document open and catalog actions are not copied into the visual Word document.",
                    PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission);
            }
            if (sourceInfo.FormFields.Any(static field => field.HasUnplacedContent) || sourceInfo.HasAcroFormXfa) {
                AddWarning(options, "PdfFormDefinitionsNotReconstructed", "Document/Forms",
                    "PDF form definitions not attached to a page and XFA content are not copied into the visual Word document.",
                    PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission);
            }
            if (sourceInfo.HasOutlines || sourceInfo.Outlines.Count > 0) {
                AddWarning(options, "PdfOutlineHierarchyNotReconstructed", "Document/Outlines",
                    "PDF outline navigation is not copied into the visual Word document.",
                    PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission);
            }
            PdfCore.PdfOptionalContentUsageSummary optionalContentUsage = source.InspectPagesForOptionalContentUsage(
                pages.Select(static page => page.PageNumber).ToArray(), token);
            if (optionalContentUsage.PagesWithUsage > 0) {
                AddWarning(options, "PdfOptionalContentGroupsFlattened", "Document/OCProperties",
                    "PDF optional-content layer controls are flattened into visual Word page images.",
                    PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission);
            }
            if (!optionalContentUsage.IsComplete && sourceInfo.HasOptionalContent) {
                AddWarning(options, "PdfOptionalContentUsageInspectionInconclusive", "Document/OCProperties",
                    "Optional-content usage could not be fully inspected for the selected PDF pages. Any selected layer controls are flattened into visual Word page images.",
                    PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Approximation);
            }
            var selectedPageNumbers = new HashSet<int>(pages.Select(static page => page.PageNumber));
            foreach (PdfCore.PdfPageInfo sourcePage in sourceInfo.Pages) {
                if (!selectedPageNumbers.Contains(sourcePage.PageNumber)) continue;
                string sourcePath = "Page " + sourcePage.PageNumber + "/";
                if (sourcePage.LinkAnnotations.Count > 0 || sourcePage.Annotations.Any(static annotation =>
                        string.Equals(annotation.Subtype, "Link", StringComparison.OrdinalIgnoreCase)))
                    AddWarning(options, "PdfLinksNotReconstructed", sourcePath + "Links",
                        "PDF interactive links are not reconstructed in the visual Word document.",
                        PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission);
                if (sourcePage.FormWidgets.Count > 0 || sourcePage.Annotations.Any(static annotation =>
                        string.Equals(annotation.Subtype, "Widget", StringComparison.OrdinalIgnoreCase)))
                    AddWarning(options, "PdfFormWidgetsNotReconstructed", sourcePath + "Forms",
                        "PDF interactive form widgets are not reconstructed in the visual Word document.",
                        PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission);
                if (sourcePage.Annotations.Any(static annotation =>
                        !string.Equals(annotation.Subtype, "Link", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(annotation.Subtype, "Widget", StringComparison.OrdinalIgnoreCase)))
                    AddWarning(options, "PdfAnnotationsNotReconstructed", sourcePath + "Annotations",
                        "PDF non-link annotations are not reconstructed in the visual Word document.",
                        PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission);
                if (sourcePage.PageActions.Count > 0)
                    AddWarning(options, "PdfPageActionsNotReconstructed", sourcePath + "Actions",
                        "PDF page actions are not reconstructed in the visual Word document.",
                        PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission);
            }
            AddWarning(options, "VisualPagesNotEditable", "Document",
                "PDF pages are embedded as images. Text, links, and forms are not editable Word objects.",
                PdfCore.PdfConversionWarningSeverity.Warning);
            for (int index = 0; index < pages.Count; index++) {
                token.ThrowIfCancellationRequested();
                PdfCore.PdfPageRenderResult page = pages[index];
                byte[] bytes = page.Bytes ?? throw new InvalidOperationException("A selected PDF page could not be rendered.");
                OfficeDrawing drawing = source.Render.Drawing(page.PageNumber);
                double width = drawing.Width, height = drawing.Height;
                // Word's supported physical page size is at most 22 inches in each dimension.
                if (!TryGetEditablePageSizeTwips(width, height, out uint widthTwips, out uint heightTwips))
                    throw new NotSupportedException("The selected PDF page is outside Word's supported physical page-size range.");
                WordSection section = index == 0 ? target.Sections[0] : target.AddSection(WordSectionBreakType.NextPage);
                section.PageSettings.Orientation = width > height ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait;
                section.PageSettings.Width = widthTwips;
                section.PageSettings.Height = heightTwips;
                section.Margins.Left = section.Margins.Right = 0;
                section.Margins.Top = section.Margins.Bottom = 0;
                section.Margins.HeaderDistance = section.Margins.FooterDistance = 0;
                WordParagraph paragraph = target.AddParagraph();
                paragraph.LineSpacingBefore = paragraph.LineSpacingAfter = 0;
                paragraph.LineSpacingPoints = 1;
                using var stream = new MemoryStream(bytes, writable: false);
                WordImage image = paragraph.InsertImage(stream, "page-" + page.PageNumber + ".png", width * 96D / 72D,
                    height * 96D / 72D, WordImageTextWrapping.InFrontOfText, "PDF page " + page.PageNumber);
                image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
                image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
                image.HorizontalPositionOffset = image.VerticalPositionOffset = 0;
                foreach (string diagnostic in page.Diagnostics)
                    AddWarning(options, "VisualPageRendering", "Page " + page.PageNumber, diagnostic, PdfCore.PdfConversionWarningSeverity.Warning);
            }
            token.ThrowIfCancellationRequested();
            return new PdfWordConversionResult(target, options.Report);
        } catch { target.Dispose(); throw; }
    }
}
