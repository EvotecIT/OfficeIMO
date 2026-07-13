using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class PdfHtmlConverterExtensions {
    /// <summary>
    /// Renders an already loaded logical PDF model as HTML and returns a machine-readable export summary.
    /// </summary>
    public static PdfHtmlConversionResult ToHtmlResult(this PdfCore.PdfLogicalDocument document, PdfHtmlSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options = (options ?? new PdfHtmlSaveOptions()).CloneForConversion();
        IReadOnlyList<PdfCore.PdfLogicalPage> pages = GetRenderPages(document, options);
        string html = options.Profile switch {
            PdfHtmlProfile.Semantic => RenderSemanticDocument(document, pages, options),
            PdfHtmlProfile.PositionedReview => RenderPositionedReviewDocument(document, pages, options),
            _ => throw new ArgumentOutOfRangeException(nameof(options.Profile), options.Profile, "Unsupported PDF HTML profile.")
        };
        return new PdfHtmlConversionResult(html, BuildExportSummary(document, pages, options, document.PageCount), options.Report);
    }

    private static PdfHtmlExportSummary BuildExportSummary(PdfCore.PdfLogicalDocument document, IReadOnlyList<PdfCore.PdfLogicalPage> pages, PdfHtmlSaveOptions options, int sourcePageCount) {
        int textBlockCount = 0;
        int headingCount = 0;
        int listItemCount = 0;
        int tableCount = 0;
        int imageCount = 0;
        int imagePlacementCount = 0;
        int imagePlaceholderCount = 0;
        int linkCount = 0;
        int renderedLinkCount = 0;
        int renderedSafeUriLinkCount = 0;
        int renderedUnsafeUriLinkCount = 0;
        int renderedInternalDestinationLinkCount = 0;
        int formWidgetCount = 0;
        var pageNumbers = new int[pages.Count];
        var formFields = new HashSet<PdfCore.PdfFormField>();
        ActionDiagnosticSummary actionSummary = BuildActionDiagnosticSummary(document, pages);

        for (int i = 0; i < pages.Count; i++) {
            PdfCore.PdfLogicalPage page = pages[i];
            pageNumbers[i] = page.PageNumber;
            textBlockCount += page.TextBlocks.Count;
            headingCount += page.Headings.Count;
            listItemCount += page.ListItems.Count;
            tableCount += page.Tables.Count;
            imageCount += page.Images.Count;
            linkCount += page.Links.Count;
            if (options.IncludeLinkAnnotations) {
                CountRenderedLinks(
                    page,
                    ref renderedLinkCount,
                    ref renderedSafeUriLinkCount,
                    ref renderedUnsafeUriLinkCount,
                    ref renderedInternalDestinationLinkCount);
            }

            formWidgetCount += page.FormWidgets.Count;
            for (int widgetIndex = 0; widgetIndex < page.FormWidgets.Count; widgetIndex++) {
                formFields.Add(page.FormWidgets[widgetIndex].Field);
            }

            for (int imageIndex = 0; imageIndex < page.Images.Count; imageIndex++) {
                PdfCore.PdfLogicalImage image = page.Images[imageIndex];
                imagePlacementCount += image.PlacementCount;
                if (options.IncludeImagePlaceholders) {
                    imagePlaceholderCount += options.Profile == PdfHtmlProfile.PositionedReview && image.HasPlacements
                        ? image.PlacementCount
                        : 1;
                }
            }
        }

        int skippedLinkCount = Math.Max(0, linkCount - renderedLinkCount);
        int outlineCount = CountOutlines(document.Outlines);
        int renderedOutlineCount = options.IncludeOutlines
            ? CountRenderedOutlines(document, pages)
            : 0;
        PdfHtmlProfileContract contract = PdfHtmlProfileContracts.Get(options.Profile);
        return new PdfHtmlExportSummary(
            options.Profile,
            contract.Id,
            pageNumbers,
            sourcePageCount,
            pages.Count,
            textBlockCount,
            headingCount,
            listItemCount,
            tableCount,
            imageCount,
            imagePlacementCount,
            imagePlaceholderCount,
            linkCount,
            renderedLinkCount,
            renderedSafeUriLinkCount,
            renderedUnsafeUriLinkCount,
            renderedInternalDestinationLinkCount,
            skippedLinkCount,
            outlineCount,
            renderedOutlineCount,
            formFields.Count,
            formWidgetCount,
            document.HasAcroFormXfa,
            document.AcroFormXfa?.PacketCount ?? 0,
            document.AcroFormXfa?.StreamCount ?? 0,
            document.AcroFormXfa?.TotalPayloadBytes ?? 0,
            actionSummary.HasOpenAction,
            actionSummary.CatalogActionCount > 0,
            actionSummary.SelectedPageActionCount > 0,
            actionSummary.SelectedAnnotationActionCount > 0,
            actionSummary.CatalogActionCount > 0 || actionSummary.SelectedPageActionCount > 0 || actionSummary.SelectedAnnotationActionCount > 0,
            actionSummary.PotentiallyUnsafeActionCount,
            actionSummary.JavaScriptActionCount,
            actionSummary.LaunchActionCount,
            actionSummary.SubmitFormActionCount,
            actionSummary.ImportDataActionCount,
            actionSummary.CatalogActionCount,
            actionSummary.PageActionCount,
            actionSummary.SelectedPageActionCount,
            actionSummary.AnnotationActionCount,
            actionSummary.SelectedAnnotationActionCount,
            options.Report.Warnings.Count,
            options.EmitDocumentShell,
            options.ImageExportMode,
            contract.FidelityContract,
            contract.UnsupportedScope);
    }

    private static void CountRenderedLinks(
        PdfCore.PdfLogicalPage page,
        ref int renderedLinkCount,
        ref int renderedSafeUriLinkCount,
        ref int renderedUnsafeUriLinkCount,
        ref int renderedInternalDestinationLinkCount) {
        for (int linkIndex = 0; linkIndex < page.Links.Count; linkIndex++) {
            PdfCore.PdfLogicalLinkAnnotation link = page.Links[linkIndex];
            if (!HasHtmlLinkTarget(link)) {
                continue;
            }

            renderedLinkCount++;
            if (link.Uri is not null) {
                if (IsSafeLinkUri(link.Uri)) {
                    renderedSafeUriLinkCount++;
                } else {
                    renderedUnsafeUriLinkCount++;
                }
            } else {
                renderedInternalDestinationLinkCount++;
            }
        }
    }

    private static ActionDiagnosticSummary BuildActionDiagnosticSummary(PdfCore.PdfLogicalDocument document, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
        int catalogActionCount = AreAllDocumentPagesSelected(document, pages) ? document.CatalogActionCount : 0;
        int selectedPageActionCount = 0;
        int selectedAnnotationActionCount = 0;
        int pageActionCount = document.PageActionCount;
        int annotationActionCount = CountAnnotationActions(document.Pages);
        var summary = new ActionDiagnosticSummary {
            HasOpenAction = HasScopedOpenAction(document.OpenAction, pages),
            CatalogActionCount = catalogActionCount,
            PageActionCount = pageActionCount,
            AnnotationActionCount = annotationActionCount
        };

        if (catalogActionCount > 0) {
            for (int i = 0; i < document.CatalogActions.Count; i++) {
                summary.Add(document.CatalogActions[i].ActionType);
            }
        }

        for (int i = 0; i < pages.Count; i++) {
            PdfCore.PdfLogicalPage page = pages[i];
            selectedPageActionCount += page.PageActionCount;
            for (int actionIndex = 0; actionIndex < page.PageActions.Count; actionIndex++) {
                summary.Add(page.PageActions[actionIndex].ActionType);
            }

            for (int annotationIndex = 0; annotationIndex < page.Annotations.Count; annotationIndex++) {
                AddAnnotationActions(page.Annotations[annotationIndex], ref selectedAnnotationActionCount, ref summary);
            }
        }

        summary.SelectedPageActionCount = selectedPageActionCount;
        summary.SelectedAnnotationActionCount = selectedAnnotationActionCount;
        return summary;
    }

    private static void AddAnnotationActions(PdfCore.PdfAnnotation annotation, ref int selectedAnnotationActionCount, ref ActionDiagnosticSummary summary) {
        if (annotation.HasAction) {
            selectedAnnotationActionCount++;
            summary.Add(annotation.ActionType);
        }

        for (int i = 0; i < annotation.AdditionalActions.Count; i++) {
            selectedAnnotationActionCount++;
            summary.Add(annotation.AdditionalActions[i].ActionType);
        }

        for (int i = 0; i < annotation.ChainedActions.Count; i++) {
            selectedAnnotationActionCount++;
            summary.Add(annotation.ChainedActions[i].ActionType);
        }
    }

    private static int CountAnnotationActions(IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
        int count = 0;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfCore.PdfLogicalPage page = pages[pageIndex];
            for (int annotationIndex = 0; annotationIndex < page.Annotations.Count; annotationIndex++) {
                PdfCore.PdfAnnotation annotation = page.Annotations[annotationIndex];
                if (annotation.HasAction) {
                    count++;
                }

                count += annotation.AdditionalActions.Count;
                count += annotation.ChainedActions.Count;
            }
        }

        return count;
    }

    private static bool AreAllDocumentPagesSelected(PdfCore.PdfLogicalDocument document, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
        if (document.PageCount == 0 || pages.Count != document.PageCount) {
            return false;
        }

        var seen = new HashSet<int>();
        for (int i = 0; i < pages.Count; i++) {
            int pageNumber = pages[i].PageNumber;
            if (pageNumber < 1 || pageNumber > document.PageCount || !seen.Add(pageNumber)) {
                return false;
            }
        }

        return seen.Count == document.PageCount;
    }

    private static bool HasScopedOpenAction(PdfCore.PdfDocumentOpenAction? openAction, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
        if (openAction is null) {
            return false;
        }

        if (!openAction.PageNumber.HasValue) {
            return true;
        }

        for (int i = 0; i < pages.Count; i++) {
            if (pages[i].PageNumber == openAction.PageNumber.Value) {
                return true;
            }
        }

        return false;
    }

    private static bool IsPotentiallyUnsafeActionType(string? actionType) =>
        string.Equals(actionType, "JavaScript", StringComparison.Ordinal) ||
        string.Equals(actionType, "Launch", StringComparison.Ordinal) ||
        string.Equals(actionType, "SubmitForm", StringComparison.Ordinal) ||
        string.Equals(actionType, "ImportData", StringComparison.Ordinal) ||
        string.Equals(actionType, "Movie", StringComparison.Ordinal) ||
        string.Equals(actionType, "RichMedia", StringComparison.Ordinal) ||
        string.Equals(actionType, "Rendition", StringComparison.Ordinal);

    private struct ActionDiagnosticSummary {
        public bool HasOpenAction { get; set; }

        public int PotentiallyUnsafeActionCount { get; private set; }

        public int JavaScriptActionCount { get; private set; }

        public int LaunchActionCount { get; private set; }

        public int SubmitFormActionCount { get; private set; }

        public int ImportDataActionCount { get; private set; }

        public int CatalogActionCount { get; set; }

        public int PageActionCount { get; set; }

        public int SelectedPageActionCount { get; set; }

        public int AnnotationActionCount { get; set; }

        public int SelectedAnnotationActionCount { get; set; }

        public void Add(string? actionType) {
            if (IsPotentiallyUnsafeActionType(actionType)) {
                PotentiallyUnsafeActionCount++;
            }

            if (string.Equals(actionType, "JavaScript", StringComparison.Ordinal)) {
                JavaScriptActionCount++;
            } else if (string.Equals(actionType, "Launch", StringComparison.Ordinal)) {
                LaunchActionCount++;
            } else if (string.Equals(actionType, "SubmitForm", StringComparison.Ordinal)) {
                SubmitFormActionCount++;
            } else if (string.Equals(actionType, "ImportData", StringComparison.Ordinal)) {
                ImportDataActionCount++;
            }
        }
    }

}
