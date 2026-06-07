namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private IReadOnlyList<PdfAnnotationChainedAction> ReadAnnotationChainedActions(PdfObject? primaryActionObject, PdfObject? additionalActionsObject) {
        var result = new List<PdfAnnotationChainedAction>();

        AddAnnotationNextActionsFromAction("A", "A.Next", primaryActionObject, result, new HashSet<int>());

        var additionalActions = ResolveDictionary(additionalActionsObject);
        if (additionalActions is not null) {
            foreach (var item in additionalActions.Items) {
                if (!string.IsNullOrEmpty(item.Key)) {
                    AddAnnotationNextActionsFromAction(item.Key, item.Key + ".Next", item.Value, result, new HashSet<int>());
                }
            }
        }

        return result.Count == 0 ? Array.Empty<PdfAnnotationChainedAction>() : result.AsReadOnly();
    }

    private void AddAnnotationNextActionsFromAction(
        string sourceName,
        string actionPath,
        PdfObject? actionObject,
        List<PdfAnnotationChainedAction> result,
        HashSet<int> visitedReferences) {
        bool enteredReference = TryEnterAnnotationActionReference(actionObject, visitedReferences);
        if (!enteredReference) {
            return;
        }

        try {
            PdfObject? resolved = ResolveObject(actionObject);
            if (resolved is PdfDictionary dictionary &&
                dictionary.Items.TryGetValue("Next", out var nextAction)) {
                AddAnnotationNextActions(sourceName, actionPath, nextAction, result, visitedReferences);
            }
        } finally {
            LeaveAnnotationActionReference(actionObject, visitedReferences);
        }
    }

    private void AddAnnotationNextActions(
        string sourceName,
        string actionPath,
        PdfObject? actionObject,
        List<PdfAnnotationChainedAction> result,
        HashSet<int> visitedReferences) {
        bool enteredReference = TryEnterAnnotationActionReference(actionObject, visitedReferences);
        if (!enteredReference) {
            return;
        }

        try {
            PdfObject? resolved = ResolveObject(actionObject);
            if (resolved is PdfArray array) {
                int activeIndex = 0;
                for (int i = 0; i < array.Items.Count; i++) {
                    int before = result.Count;
                    AddAnnotationChainedAction(sourceName, actionPath + "." + activeIndex.ToString(System.Globalization.CultureInfo.InvariantCulture), array.Items[i], result, visitedReferences);
                    if (result.Count > before) {
                        activeIndex++;
                    }
                }

                return;
            }

            AddAnnotationChainedAction(sourceName, actionPath, resolved, result, visitedReferences);
        } finally {
            LeaveAnnotationActionReference(actionObject, visitedReferences);
        }
    }

    private void AddAnnotationChainedAction(
        string sourceName,
        string actionPath,
        PdfObject? actionObject,
        List<PdfAnnotationChainedAction> result,
        HashSet<int> visitedReferences) {
        bool enteredReference = TryEnterAnnotationActionReference(actionObject, visitedReferences);
        if (!enteredReference) {
            return;
        }

        try {
            PdfObject? resolved = ResolveObject(actionObject);
            if (resolved is not PdfDictionary dictionary) {
                return;
            }

            string? actionType = TryReadActionType(dictionary);
            if (!string.IsNullOrEmpty(actionType)) {
                result.Add(new PdfAnnotationChainedAction(sourceName, actionPath, actionType!));
            }

            if (dictionary.Items.TryGetValue("Next", out var nextAction)) {
                AddAnnotationNextActions(sourceName, actionPath + ".Next", nextAction, result, visitedReferences);
            }
        } finally {
            LeaveAnnotationActionReference(actionObject, visitedReferences);
        }
    }

    private static bool TryEnterAnnotationActionReference(PdfObject? actionObject, HashSet<int> visitedReferences) {
        return actionObject is not PdfReference reference || visitedReferences.Add(reference.ObjectNumber);
    }

    private static void LeaveAnnotationActionReference(PdfObject? actionObject, HashSet<int> visitedReferences) {
        if (actionObject is PdfReference reference) {
            visitedReferences.Remove(reference.ObjectNumber);
        }
    }
}
