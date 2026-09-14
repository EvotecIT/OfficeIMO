using System.Collections.Generic;
using System.Linq;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class PdfHtmlConverterExtensions {
    /// <summary>Renders an opened PDF as HTML and returns a machine-readable export summary.</summary>
    public static PdfHtmlConversionResult ToHtmlResult(this PdfCore.PdfDocument document, PdfToHtmlOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadForHtml(document, options, cancellationToken)
            .ToHtmlResult(CreateRenderOptionsAfterPreselection(options, document), cancellationToken);
    }

    /// <summary>
    /// Renders an already loaded logical PDF model as HTML and returns a machine-readable export summary.
    /// </summary>
    public static PdfHtmlConversionResult ToHtmlResult(this PdfCore.PdfDocumentReadResult document, PdfToHtmlOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options = (options ?? new PdfToHtmlOptions()).CloneForConversion();
        options.CancellationToken = cancellationToken;
        options.CancellationToken.ThrowIfCancellationRequested();
        options.Validate();
        IReadOnlyList<PdfCore.PdfLogicalPage> pages = GetRenderPages(document, options);
        string html;
        try {
            html = options.Profile switch {
                PdfHtmlProfile.Semantic => RenderSemanticDocument(document, pages, options),
                PdfHtmlProfile.PositionedReview => RenderPositionedReviewDocument(document, pages, options),
                _ => throw new ArgumentOutOfRangeException(nameof(options.Profile), options.Profile, "Unsupported PDF HTML profile.")
            };
        } catch (Exception exception) when (
            options.MaximumOutputCharacters.HasValue &&
            IsOutputBuilderCapacityException(exception)) {
            throw new InvalidOperationException(
                $"Generated HTML exceeded the configured {options.MaximumOutputCharacters.Value:N0}-character output limit while it was being rendered.",
                exception);
        }
        ReportProfileFidelity(document, pages, options);
        return new PdfHtmlConversionResult(html, BuildExportSummary(document, pages, options, document.SourcePageCount), options.Report);
    }

    private static void ReportProfileFidelity(
        PdfCore.PdfDocumentReadResult document,
        IReadOnlyList<PdfCore.PdfLogicalPage> pages,
        PdfToHtmlOptions options) {
        int textBlockCount = 0;
        int tableCount = 0;
        int imageCount = 0;
        int linkCount = 0;
        int skippedLinkCount = 0;
        int formWidgetCount = 0;
        int unrepresentedVectorCount = 0;
        int annotationCount = 0;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            PdfCore.PdfLogicalPage page = pages[pageIndex];
            textBlockCount += page.TextBlocks.Count;
            tableCount += page.Tables.Count;
            imageCount += page.Images.Count(image =>
                PdfCore.PdfImagePlacementImportPolicy.HasVisiblePlacement(page, image));
            linkCount += page.Links.Count;
            formWidgetCount += page.FormWidgets.Count;
            unrepresentedVectorCount += page.UnrepresentedVectorPrimitiveCount;
            annotationCount += page.Annotations.Count(static annotation =>
                !string.Equals(annotation.Subtype, "Link", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(annotation.Subtype, "Widget", StringComparison.OrdinalIgnoreCase));
            if (options.IncludeLinkAnnotations) {
                skippedLinkCount += page.Links.Count(static link => !HasHtmlLinkTarget(link));
            }
        }

        if (options.Profile == PdfHtmlProfile.Semantic &&
            textBlockCount + tableCount + imageCount + unrepresentedVectorCount > 0) {
            AddWarning(
                options,
                "PdfSemanticLayoutReflowed",
                "Semantic HTML reconstructs readable document structure; fixed PDF coordinates, pagination, and authoring layout are not preserved.",
                PdfCore.PdfConversionWarningSeverity.Information,
                OfficeConversionLossKind.Approximation);
        }

        if (options.Profile == PdfHtmlProfile.Semantic && unrepresentedVectorCount > 0) {
            AddWarning(
                options,
                "PdfVectorAppearanceOmitted",
                unrepresentedVectorCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " vector primitives were not represented by HTML content or detected table structure.",
                PdfCore.PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.Omission);
        }

        if (!options.IncludeImagePlaceholders && imageCount > 0) {
            AddWarning(
                options,
                "PdfImagesOmitted",
                imageCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " PDF images were omitted because image output is disabled.",
                PdfCore.PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.Omission);
        } else if (options.IncludeImagePlaceholders &&
                   options.ImageExportMode == PdfHtmlImageExportMode.PlaceholderOnly &&
                   imageCount > 0) {
            AddWarning(
                options,
                "PdfImagePixelsOmitted",
                imageCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " PDF images were represented by readable placeholders without their source pixels.",
                PdfCore.PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.Omission);
        }

        skippedLinkCount += options.IncludeLinkAnnotations ? 0 : linkCount;
        if (skippedLinkCount > 0) {
            AddWarning(
                options,
                "PdfLinkAnnotationsOmitted",
                skippedLinkCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " PDF link annotations were omitted because link output is disabled or the target is unsupported.",
                PdfCore.PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.Omission);
        }

        if (formWidgetCount > 0) {
            if (options.IncludeFormWidgets) {
                AddWarning(
                    options,
                    "PdfFormWidgetsFlattened",
                    formWidgetCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                    " PDF form widgets were represented as static text; field editing and interactive behavior were not preserved.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    OfficeConversionLossKind.Approximation);
            } else {
                AddWarning(
                    options,
                    "PdfFormWidgetsOmitted",
                    formWidgetCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                    " PDF form widgets were omitted because form output is disabled.",
                    PdfCore.PdfConversionWarningSeverity.Warning,
                    OfficeConversionLossKind.Omission);
            }
        }

        if (annotationCount > 0) {
            AddWarning(
                options,
                "PdfAnnotationsOmitted",
                annotationCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " non-link PDF annotations are not represented by the HTML profiles.",
                PdfCore.PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.Omission);
        }

        if (document.OptionalContentGroupCount > 0) {
            AddWarning(
                options,
                "PdfOptionalContentGroupsFlattened",
                document.OptionalContentGroupCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " PDF optional-content groups were flattened to the current visible projection; layer controls and alternate visibility states were not preserved.",
                PdfCore.PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.Approximation);
        }

        int outlineCount = CountOutlines(document.Outlines);
        int renderedOutlineCount = options.IncludeOutlines
            ? CountRenderedOutlines(document, pages)
            : 0;
        int omittedOutlineCount = Math.Max(0, outlineCount - renderedOutlineCount);
        if (omittedOutlineCount > 0) {
            AddWarning(
                options,
                "PdfOutlinesOmitted",
                omittedOutlineCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " PDF outline entries were omitted because outline output is disabled or their destination is outside the selected pages.",
                PdfCore.PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.Omission);
        }

        ActionDiagnosticSummary actionSummary = BuildActionDiagnosticSummary(document, pages);
        int omittedDocumentActionCount = actionSummary.CatalogActionCount +
            actionSummary.SelectedPageActionCount +
            CountOmittedAnnotationActions(pages, options.IncludeLinkAnnotations) +
            (actionSummary.HasOpenAction ? 1 : 0);
        if (omittedDocumentActionCount > 0) {
            AddWarning(
                options,
                "PdfDocumentActionsOmitted",
                omittedDocumentActionCount.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                " scoped PDF open, catalog, page, or annotation actions were stripped from HTML output.",
                PdfCore.PdfConversionWarningSeverity.Warning,
                OfficeConversionLossKind.Omission);
        }
    }

    private static bool IsOutputBuilderCapacityException(Exception exception) =>
        exception is PdfHtmlOutputCapacityException ||
        exception is ArgumentOutOfRangeException argumentException &&
        (string.Equals(argumentException.ParamName, "valueCount", StringComparison.Ordinal) ||
         string.Equals(argumentException.ParamName, "requiredLength", StringComparison.Ordinal) ||
         string.Equals(argumentException.ParamName, "repeatCount", StringComparison.Ordinal)) &&
        argumentException.StackTrace?.IndexOf("System.Text.StringBuilder", StringComparison.Ordinal) >= 0;

    private static PdfHtmlExportSummary BuildExportSummary(PdfCore.PdfDocumentReadResult document, IReadOnlyList<PdfCore.PdfLogicalPage> pages, PdfToHtmlOptions options, int sourcePageCount) {
        int textBlockCount = 0;
        int headingCount = 0;
        int listItemCount = 0;
        int tableCount = 0;
        int imageCount = 0;
        int imagePlacementCount = 0;
        int imagePlaceholderCount = options.EmittedImagePlaceholderCount;
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
            options.CancellationToken.ThrowIfCancellationRequested();
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
            options.EmitDocumentShell && options.IncludeDefaultStyles,
            options.Theme,
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

    private static ActionDiagnosticSummary BuildActionDiagnosticSummary(PdfCore.PdfDocumentReadResult document, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
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

    private static int CountOmittedAnnotationActions(
        IReadOnlyList<PdfCore.PdfLogicalPage> pages,
        bool includeLinkAnnotations) {
        int omittedCount = 0;
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            PdfCore.PdfLogicalPage page = pages[pageIndex];
            var representedLinks = new HashSet<int>();
            for (int annotationIndex = 0; annotationIndex < page.Annotations.Count; annotationIndex++) {
                PdfCore.PdfAnnotation annotation = page.Annotations[annotationIndex];
                omittedCount += annotation.AdditionalActions.Count;
                omittedCount += annotation.ChainedActions.Count;
                if (!annotation.HasAction) {
                    continue;
                }

                if (!includeLinkAnnotations ||
                    !TryMatchRepresentedPrimaryLinkAction(page, annotation, representedLinks)) {
                    omittedCount++;
                }
            }
        }

        return omittedCount;
    }

    private static bool TryMatchRepresentedPrimaryLinkAction(
        PdfCore.PdfLogicalPage page,
        PdfCore.PdfAnnotation annotation,
        HashSet<int> representedLinks) {
        if (!string.Equals(annotation.Subtype, "Link", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        for (int linkIndex = 0; linkIndex < page.Links.Count; linkIndex++) {
            if (representedLinks.Contains(linkIndex)) {
                continue;
            }

            PdfCore.PdfLogicalLinkAnnotation link = page.Links[linkIndex];
            if (!HasSameRectangle(annotation, link) || !IsPrimaryLinkActionRepresented(annotation, link)) {
                continue;
            }

            representedLinks.Add(linkIndex);
            return true;
        }

        return false;
    }

    private static bool IsPrimaryLinkActionRepresented(
        PdfCore.PdfAnnotation annotation,
        PdfCore.PdfLogicalLinkAnnotation link) {
        if (string.Equals(annotation.ActionType, "URI", StringComparison.OrdinalIgnoreCase)) {
            return link.Uri is not null && IsSafeLinkUri(link.Uri);
        }

        if (string.Equals(annotation.ActionType, "GoTo", StringComparison.OrdinalIgnoreCase)) {
            return !string.IsNullOrWhiteSpace(link.DestinationName) || link.DestinationPageNumber.HasValue;
        }

        return false;
    }

    private static bool HasSameRectangle(
        PdfCore.PdfAnnotation annotation,
        PdfCore.PdfLogicalLinkAnnotation link) {
        const double tolerance = 0.001D;
        return Math.Abs(annotation.X1 - link.X1) <= tolerance &&
            Math.Abs(annotation.Y1 - link.Y1) <= tolerance &&
            Math.Abs(annotation.X2 - link.X2) <= tolerance &&
            Math.Abs(annotation.Y2 - link.Y2) <= tolerance;
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

    private static bool AreAllDocumentPagesSelected(PdfCore.PdfDocumentReadResult document, IReadOnlyList<PdfCore.PdfLogicalPage> pages) {
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
