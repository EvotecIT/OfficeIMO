namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private IReadOnlyList<PdfAnnotationChainedAction> ReadAnnotationChainedActions(PdfObject? primaryActionObject, PdfObject? additionalActionsObject) {
        var result = new List<PdfAnnotationChainedAction>();
        var visitedReferences = new HashSet<int>();

        AddAnnotationNextActionsFromAction("A", "A.Next", primaryActionObject, result, visitedReferences);

        var additionalActions = ResolveDictionary(additionalActionsObject);
        if (additionalActions is not null) {
            foreach (var item in additionalActions.Items) {
                if (!string.IsNullOrEmpty(item.Key)) {
                    AddAnnotationNextActionsFromAction(item.Key, item.Key + ".Next", item.Value, result, visitedReferences);
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
        PdfObject? resolved = ResolveObject(actionObject);
        if (actionObject is PdfReference reference && !visitedReferences.Add(reference.ObjectNumber)) {
            return;
        }

        if (resolved is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("Next", out var nextAction)) {
            AddAnnotationNextActions(sourceName, actionPath, nextAction, result, visitedReferences);
        }
    }

    private void AddAnnotationNextActions(
        string sourceName,
        string actionPath,
        PdfObject? actionObject,
        List<PdfAnnotationChainedAction> result,
        HashSet<int> visitedReferences) {
        PdfObject? resolved = ResolveObject(actionObject);
        if (actionObject is PdfReference reference && !visitedReferences.Add(reference.ObjectNumber)) {
            return;
        }

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
    }

    private void AddAnnotationChainedAction(
        string sourceName,
        string actionPath,
        PdfObject? actionObject,
        List<PdfAnnotationChainedAction> result,
        HashSet<int> visitedReferences) {
        PdfObject? resolved = ResolveObject(actionObject);
        if (actionObject is PdfReference reference && !visitedReferences.Add(reference.ObjectNumber)) {
            return;
        }

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
    }
}
