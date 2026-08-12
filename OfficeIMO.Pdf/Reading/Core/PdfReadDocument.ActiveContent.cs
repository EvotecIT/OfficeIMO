namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    internal bool HasOnlyWidgetOwnedActiveContent() {
        if (_acroFormXfa is not null) return false;

        var widgetObjectNumbers = new HashSet<int>();
        for (int fieldIndex = 0; fieldIndex < _formFields.Count; fieldIndex++) {
            IReadOnlyList<PdfFormWidget> widgets = _formFields[fieldIndex].Widgets;
            for (int widgetIndex = 0; widgetIndex < widgets.Count; widgetIndex++) {
                PdfFormWidget widget = widgets[widgetIndex];
                if (widget.HasActions && widget.ObjectNumber.HasValue) {
                    widgetObjectNumbers.Add(widget.ObjectNumber.Value);
                }
            }
        }

        if (widgetObjectNumbers.Count == 0) return false;
        PdfDictionary? catalog = FindCatalog();
        return catalog is not null && !ContainsActiveContentOutsideWidgets(
            catalog,
            widgetObjectNumbers);
    }

    private bool ContainsActiveContentOutsideWidgets(
        PdfObject value,
        HashSet<int> widgetObjectNumbers) {
        var visitedReferences = new HashSet<(int ObjectNumber, int Generation)>();
        var pending = new Stack<(PdfObject Value, int Depth, bool WidgetRoot)>();
        pending.Push((value, 0, false));
        while (pending.Count > 0) {
            (PdfObject current, int depth, bool widgetRoot) = pending.Pop();
            if (depth > _options.Limits.MaxObjectNestingDepth) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectNestingDepth, _options.Limits.MaxObjectNestingDepth, depth);
            }

            if (current is PdfReference reference) {
                if (!visitedReferences.Add((reference.ObjectNumber, reference.Generation)) ||
                    !PdfObjectLookup.TryGet(_objects, reference, out PdfIndirectObject? indirect)) {
                    continue;
                }
                pending.Push((indirect.Value, depth + 1, widgetObjectNumbers.Contains(reference.ObjectNumber)));
                continue;
            }
            if (current is PdfStream stream) {
                pending.Push((stream.Dictionary, depth + 1, widgetRoot));
                continue;
            }
            if (current is PdfArray array) {
                for (int index = array.Items.Count - 1; index >= 0; index--) pending.Push((array.Items[index], depth + 1, false));
                continue;
            }
            if (current is not PdfDictionary dictionary) continue;
            for (int index = 0; index < PdfActiveContentPolicy.MarkerNames.Length; index++) {
                string marker = PdfActiveContentPolicy.MarkerNames[index];
                if ((!widgetRoot || !string.Equals(marker, "AA", StringComparison.Ordinal)) && dictionary.Items.ContainsKey(marker)) return true;
            }
            foreach (KeyValuePair<string, PdfObject> item in dictionary.Items) {
                if (widgetRoot && (string.Equals(item.Key, "A", StringComparison.Ordinal) || string.Equals(item.Key, "AA", StringComparison.Ordinal))) continue;
                if (item.Value is PdfName name) {
                    for (int index = 0; index < PdfActiveContentPolicy.MarkerNames.Length; index++) {
                        if (string.Equals(name.Name, PdfActiveContentPolicy.MarkerNames[index], StringComparison.Ordinal)) return true;
                    }
                }
                pending.Push((item.Value, depth + 1, false));
            }
        }
        return false;
    }
}
