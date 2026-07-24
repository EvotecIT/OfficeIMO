using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

internal static partial class PdfReaderAdapter {
    private static IReadOnlyList<ReaderActionSummary>? BuildActions(PdfLogicalDocument document, IReadOnlyList<PdfLogicalPage> selectedPages, PdfLogicalPage? page, bool includeDocumentActions) {
        IReadOnlyList<PdfLogicalPage> scope = page is null ? selectedPages : new[] { page };
        PdfDocumentOpenAction? scopedOpenAction = includeDocumentActions ? GetScopedOpenAction(document.OpenAction, scope) : null;
        IReadOnlyList<PdfCatalogAction> catalogActions = includeDocumentActions
            ? GetScopedCatalogActions(document)
            : Array.Empty<PdfCatalogAction>();
        bool hasOpenAction = scopedOpenAction is not null;
        bool hasCatalogActions = catalogActions.Count > 0;
        int selectedPageActionCount = CountPageActions(scope);
        int selectedAnnotationActionCount = CountAnnotationActions(scope);
        if (!hasOpenAction && !hasCatalogActions && selectedPageActionCount == 0 && selectedAnnotationActionCount == 0) {
            return null;
        }

        var actions = new List<ReaderActionSummary>();
        if (scopedOpenAction is not null) {
            actions.Add(BuildOpenAction(scopedOpenAction));
        }

        for (int i = 0; i < catalogActions.Count; i++) {
            actions.Add(BuildCatalogAction(catalogActions[i]));
        }

        for (int i = 0; i < scope.Count; i++) {
            PdfLogicalPage logicalPage = scope[i];
            for (int j = 0; j < logicalPage.PageActions.Count; j++) {
                actions.Add(BuildPageAction(logicalPage.PageActions[j]));
            }

            for (int j = 0; j < logicalPage.Annotations.Count; j++) {
                AddAnnotationActions(logicalPage.Annotations[j], actions);
            }
        }

        return actions.Count == 0 ? null : actions.AsReadOnly();
    }

    private static IReadOnlyList<PdfCatalogAction> GetScopedCatalogActions(PdfLogicalDocument document) =>
        document.CatalogActions;

    private static PdfDocumentOpenAction? GetScopedOpenAction(PdfDocumentOpenAction? openAction, IReadOnlyList<PdfLogicalPage> scope) {
        if (openAction is null) {
            return null;
        }

        if (!openAction.PageNumber.HasValue) {
            return openAction;
        }

        int pageNumber = openAction.PageNumber.Value;
        for (int i = 0; i < scope.Count; i++) {
            if (scope[i].PageNumber == pageNumber) {
                return openAction;
            }
        }

        return null;
    }

    private static ReaderActionSummary BuildOpenAction(PdfDocumentOpenAction action) {
        return new ReaderActionSummary {
            Scope = ReaderActionScope.DocumentOpen,
            ActionType = action.ActionType,
            Source = "OpenAction",
            IsPotentiallyUnsafe = IsPotentiallyUnsafeActionType(action.ActionType),
            DestinationPageNumber = action.PageNumber,
            DestinationMode = action.DestinationMode?.ToString(),
            DestinationTop = action.DestinationTop,
            DestinationLeft = action.DestinationLeft,
            DestinationBottom = action.DestinationBottom,
            DestinationRight = action.DestinationRight
        };
    }

    private static ReaderActionSummary BuildCatalogAction(PdfCatalogAction action) {
        return new ReaderActionSummary {
            Scope = ReaderActionScope.Catalog,
            ActionType = action.ActionType,
            Source = action.Source,
            Name = action.Name,
            TriggerName = action.TriggerName,
            ActionPath = action.ActionPath,
            IsChainedAction = action.IsChainedAction,
            IsPotentiallyUnsafe = IsPotentiallyUnsafeActionType(action.ActionType)
        };
    }

    private static ReaderActionSummary BuildPageAction(PdfPageAction action) {
        return new ReaderActionSummary {
            Scope = ReaderActionScope.Page,
            ActionType = action.ActionType,
            Source = "Page/AA",
            TriggerName = action.TriggerName,
            ActionPath = action.ActionPath,
            PageNumber = action.PageNumber,
            IsChainedAction = action.IsChainedAction,
            IsPotentiallyUnsafe = IsPotentiallyUnsafeActionType(action.ActionType)
        };
    }

    private static void AddAnnotationActions(PdfAnnotation annotation, List<ReaderActionSummary> actions) {
        if (annotation.HasAction) {
            actions.Add(new ReaderActionSummary {
                Scope = ReaderActionScope.Annotation,
                ActionType = annotation.ActionType!,
                Source = "Annotation/A",
                Name = annotation.Subtype,
                ActionPath = "A",
                PageNumber = annotation.PageNumber,
                IsChainedAction = false,
                IsPotentiallyUnsafe = IsPotentiallyUnsafeActionType(annotation.ActionType)
            });
        }

        for (int i = 0; i < annotation.AdditionalActions.Count; i++) {
            PdfAnnotationAdditionalAction action = annotation.AdditionalActions[i];
            actions.Add(new ReaderActionSummary {
                Scope = ReaderActionScope.Annotation,
                ActionType = action.ActionType,
                Source = "Annotation/AA",
                Name = annotation.Subtype,
                TriggerName = action.TriggerName,
                ActionPath = "AA." + action.TriggerName,
                PageNumber = annotation.PageNumber,
                IsChainedAction = false,
                IsPotentiallyUnsafe = IsPotentiallyUnsafeActionType(action.ActionType)
            });
        }

        for (int i = 0; i < annotation.ChainedActions.Count; i++) {
            PdfAnnotationChainedAction action = annotation.ChainedActions[i];
            actions.Add(new ReaderActionSummary {
                Scope = ReaderActionScope.Annotation,
                ActionType = action.ActionType,
                Source = "Annotation/Next",
                Name = annotation.Subtype,
                TriggerName = action.SourceName,
                ActionPath = action.ActionPath,
                PageNumber = annotation.PageNumber,
                IsChainedAction = true,
                IsPotentiallyUnsafe = IsPotentiallyUnsafeActionType(action.ActionType)
            });
        }
    }

    private static bool IsPotentiallyUnsafeActionType(string? actionType) =>
        string.Equals(actionType, "JavaScript", StringComparison.Ordinal) ||
        string.Equals(actionType, "Launch", StringComparison.Ordinal) ||
        string.Equals(actionType, "SubmitForm", StringComparison.Ordinal) ||
        string.Equals(actionType, "ImportData", StringComparison.Ordinal) ||
        string.Equals(actionType, "RichMedia", StringComparison.Ordinal) ||
        string.Equals(actionType, "Movie", StringComparison.Ordinal) ||
        string.Equals(actionType, "Rendition", StringComparison.Ordinal);
}
