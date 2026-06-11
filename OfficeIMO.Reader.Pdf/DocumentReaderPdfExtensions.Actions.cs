using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

public static partial class DocumentReaderPdfExtensions {
    private static IReadOnlyList<ReaderActionSummary>? BuildActions(PdfLogicalDocument document, IReadOnlyList<PdfLogicalPage> selectedPages, PdfLogicalPage? page) {
        bool hasOpenAction = document.OpenAction is not null;
        bool hasCatalogActions = document.CatalogActions.Count > 0;
        IReadOnlyList<PdfLogicalPage> scope = page is null ? selectedPages : new[] { page };
        int selectedPageActionCount = CountPageActions(scope);
        int selectedAnnotationActionCount = CountAnnotationActions(scope);
        if (!hasOpenAction && !hasCatalogActions && selectedPageActionCount == 0 && selectedAnnotationActionCount == 0) {
            return null;
        }

        var actions = new List<ReaderActionSummary>();
        if (document.OpenAction is not null) {
            actions.Add(BuildOpenAction(document.OpenAction));
        }

        for (int i = 0; i < document.CatalogActions.Count; i++) {
            actions.Add(BuildCatalogAction(document.CatalogActions[i]));
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

    private static ReaderActionSummary BuildOpenAction(PdfDocumentOpenAction action) {
        return new ReaderActionSummary {
            Scope = ReaderActionScope.DocumentOpen,
            ActionType = action.ActionType,
            Source = "OpenAction",
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
            TriggerName = action.TriggerName
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
            IsChainedAction = action.IsChainedAction
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
                IsChainedAction = false
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
                IsChainedAction = false
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
                IsChainedAction = true
            });
        }
    }
}
