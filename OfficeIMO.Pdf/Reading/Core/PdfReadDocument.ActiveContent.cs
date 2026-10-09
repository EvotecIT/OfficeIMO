namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    internal bool HasOnlyFormOwnedActiveContent() {
        if (_acroFormXfa is not null) return false;

        var formObjectNumbers = new HashSet<int>();
        for (int fieldIndex = 0; fieldIndex < _formFields.Count; fieldIndex++) {
            PdfFormField field = _formFields[fieldIndex];
            // A terminal field with widget children has its own action surface. A non-widget
            // annotation carrying /FT is not thereby a legitimate field dictionary.
            if (field.Actions.Count > 0 && field.Widgets.Count > 0 && field.ObjectNumber is int number &&
                _objects.TryGetValue(number, out var indirect) && indirect.Value is PdfDictionary dictionary &&
                !dictionary.Items.ContainsKey("Subtype") && !dictionary.Items.ContainsKey("Type")) {
                formObjectNumbers.Add(number);
            }
            IReadOnlyList<PdfFormWidget> widgets = field.Widgets;
            for (int widgetIndex = 0; widgetIndex < widgets.Count; widgetIndex++) {
                PdfFormWidget widget = widgets[widgetIndex];
                if (widget.HasActions && widget.ObjectNumber.HasValue) {
                    formObjectNumbers.Add(widget.ObjectNumber.Value);
                }
            }
        }

        if (formObjectNumbers.Count == 0) return false;
        PdfDictionary? catalog = FindCatalog();
        return catalog is not null && !ContainsActiveContentOutsideForms(
            catalog,
            formObjectNumbers);
    }

    private bool ContainsActiveContentOutsideForms(
        PdfObject value,
        HashSet<int> formObjectNumbers) {
        var visitedReferences = new HashSet<(int ObjectNumber, int Generation)>();
        var pending = new Stack<(PdfObject Value, int Depth, bool FormRoot)>();
        pending.Push((value, 0, false));
        while (pending.Count > 0) {
            (PdfObject current, int depth, bool formRoot) = pending.Pop();
            if (depth > _options.Limits.MaxObjectNestingDepth) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ObjectNestingDepth, _options.Limits.MaxObjectNestingDepth, depth);
            }

            if (current is PdfReference reference) {
                if (!visitedReferences.Add((reference.ObjectNumber, reference.Generation)) ||
                    !PdfObjectLookup.TryGet(_objects, reference, out PdfIndirectObject? indirect)) {
                    continue;
                }
                pending.Push((indirect.Value, depth + 1, formObjectNumbers.Contains(reference.ObjectNumber)));
                continue;
            }
            if (current is PdfStream stream) {
                pending.Push((stream.Dictionary, depth + 1, formRoot));
                continue;
            }
            if (current is PdfName name) {
                for (int index = 0; index < PdfActiveContentPolicy.MarkerNames.Length; index++) {
                    if (string.Equals(name.Name, PdfActiveContentPolicy.MarkerNames[index], StringComparison.Ordinal)) return true;
                }
                continue;
            }
            if (current is PdfArray array) {
                for (int index = array.Items.Count - 1; index >= 0; index--) pending.Push((array.Items[index], depth + 1, false));
                continue;
            }
            if (current is not PdfDictionary dictionary) continue;
            for (int index = 0; index < PdfActiveContentPolicy.MarkerNames.Length; index++) {
                string marker = PdfActiveContentPolicy.MarkerNames[index];
                if ((!formRoot || !string.Equals(marker, "AA", StringComparison.Ordinal)) && dictionary.Items.ContainsKey(marker)) return true;
            }
            foreach (KeyValuePair<string, PdfObject> item in dictionary.Items) {
                if (formRoot && (string.Equals(item.Key, "A", StringComparison.Ordinal) || string.Equals(item.Key, "AA", StringComparison.Ordinal))) continue;
                pending.Push((item.Value, depth + 1, false));
            }
        }
        return false;
    }
}
