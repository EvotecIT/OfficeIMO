namespace OfficeIMO.Pdf;

internal static partial class PdfAcroFormEditor {
    private static void ValidatePlannedWidgetJavaScriptBudget(
        PdfReadDocument source,
        IReadOnlyList<PdfAcroFormEditSession.EditCommand> commands) {
        var contributions = new Dictionary<string, WidgetJavaScriptContribution>(StringComparer.Ordinal);
        for (int fieldIndex = 0; fieldIndex < source.UncheckedFormFields.Count; fieldIndex++) {
            PdfFormField field = source.UncheckedFormFields[fieldIndex];
            int count = 0;
            long bytes = 0L;
            int totalCount = 0;
            for (int widgetIndex = 0; widgetIndex < field.Widgets.Count; widgetIndex++) {
                IReadOnlyList<PdfFormWidgetAction> actions = field.Widgets[widgetIndex].Actions;
                for (int actionIndex = 0; actionIndex < actions.Count; actionIndex++) {
                    PdfFormWidgetAction action = actions[actionIndex];
                    totalCount = checked(totalCount + 1);
                    if (!action.IsJavaScript) continue;
                    count = checked(count + 1);
                    bytes = checked(bytes + action.JavaScriptSourceBytes);
                }
            }
            if (totalCount > 0) {
                string key = string.IsNullOrEmpty(field.Name)
                    ? "\0" + fieldIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : field.Name!;
                WidgetJavaScriptContribution contribution = new WidgetJavaScriptContribution(count, bytes, totalCount);
                if (contributions.TryGetValue(key, out WidgetJavaScriptContribution existing)) {
                    contribution = new WidgetJavaScriptContribution(
                        checked(existing.Count + contribution.Count),
                        checked(existing.Bytes + contribution.Bytes),
                        checked(existing.TotalCount + contribution.TotalCount));
                }
                contributions[key] = contribution;
            }
        }

        ValidateWidgetJavaScriptBudget(contributions, source.ReadOptions.Limits);
        for (int commandIndex = 0; commandIndex < commands.Count; commandIndex++) {
            PdfAcroFormEditSession.EditCommand command = commands[commandIndex];
            switch (command.Kind) {
                case PdfAcroFormEditSession.EditKind.Create:
                    if (command.EncodedJavaScript is not null) {
                        int count = command.Options!.Kind == PdfFormFieldCreationKind.RadioButtonGroup
                            ? command.Options.ChoiceOptions.Count
                            : 1;
                        contributions[command.Options.Name] = new WidgetJavaScriptContribution(
                            count,
                            checked(command.EncodedJavaScript.LongLength * count),
                            count);
                    }
                    break;
                case PdfAcroFormEditSession.EditKind.Remove:
                    RemoveWidgetJavaScriptContribution(contributions, command.Name!);
                    break;
                case PdfAcroFormEditSession.EditKind.Rename:
                    RenameWidgetJavaScriptContribution(contributions, command.Name!, command.Value!);
                    break;
            }
            ValidateWidgetJavaScriptBudget(contributions, source.ReadOptions.Limits);
        }
    }

    private static void ValidateWidgetJavaScriptBudget(
        IReadOnlyDictionary<string, WidgetJavaScriptContribution> contributions,
        PdfReadLimits limits) {
        int count = 0;
        long bytes = 0L;
        int totalCount = 0;
        foreach (WidgetJavaScriptContribution contribution in contributions.Values) {
            count = checked(count + contribution.Count);
            bytes = checked(bytes + contribution.Bytes);
            totalCount = checked(totalCount + contribution.TotalCount);
        }
        if (count > limits.MaxJavaScripts) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.JavaScripts, limits.MaxJavaScripts, count);
        }
        if (bytes > limits.MaxTotalJavaScriptBytes) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.JavaScriptBytes, limits.MaxTotalJavaScriptBytes, bytes);
        }
        if (totalCount > limits.MaxWidgetActions) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.WidgetActions, limits.MaxWidgetActions, totalCount);
        }
    }

    private static void RemoveWidgetJavaScriptContribution(
        Dictionary<string, WidgetJavaScriptContribution> contributions,
        string name) {
        string descendantPrefix = name + ".";
        foreach (string fieldName in contributions.Keys
                     .Where(candidate => string.Equals(candidate, name, StringComparison.Ordinal) || candidate.StartsWith(descendantPrefix, StringComparison.Ordinal))
                     .ToArray()) {
            contributions.Remove(fieldName);
        }
    }

    private static void RenameWidgetJavaScriptContribution(
        Dictionary<string, WidgetJavaScriptContribution> contributions,
        string name,
        string newName) {
        string descendantPrefix = name + ".";
        foreach (KeyValuePair<string, WidgetJavaScriptContribution> contribution in contributions
                     .Where(item => string.Equals(item.Key, name, StringComparison.Ordinal) || item.Key.StartsWith(descendantPrefix, StringComparison.Ordinal))
                     .ToArray()) {
            contributions.Remove(contribution.Key);
            contributions[newName + contribution.Key.Remove(0, name.Length)] = contribution.Value;
        }
    }

    private readonly struct WidgetJavaScriptContribution {
        internal WidgetJavaScriptContribution(int count, long bytes, int totalCount) { Count = count; Bytes = bytes; TotalCount = totalCount; }
        internal int Count { get; }
        internal long Bytes { get; }
        internal int TotalCount { get; }
    }
}
