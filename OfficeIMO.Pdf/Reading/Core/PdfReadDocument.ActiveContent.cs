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
            widgetObjectNumbers,
            new HashSet<int>());
    }

    private bool ContainsActiveContentOutsideWidgets(
        PdfObject value,
        HashSet<int> widgetObjectNumbers,
        HashSet<int> visitedReferences) {
        if (value is PdfReference reference) {
            if (widgetObjectNumbers.Contains(reference.ObjectNumber)) return false;
            if (!visitedReferences.Add(reference.ObjectNumber) ||
                !PdfObjectLookup.TryGet(_objects, reference, out PdfIndirectObject? indirect)) {
                return false;
            }

            return ContainsActiveContentOutsideWidgets(indirect.Value, widgetObjectNumbers, visitedReferences);
        }

        if (value is PdfStream stream) {
            return ContainsActiveContentOutsideWidgets(stream.Dictionary, widgetObjectNumbers, visitedReferences);
        }

        if (value is PdfArray array) {
            for (int index = 0; index < array.Items.Count; index++) {
                if (ContainsActiveContentOutsideWidgets(array.Items[index], widgetObjectNumbers, visitedReferences)) {
                    return true;
                }
            }

            return false;
        }

        if (value is not PdfDictionary dictionary) return false;
        for (int index = 0; index < PdfActiveContentPolicy.MarkerNames.Length; index++) {
            string marker = PdfActiveContentPolicy.MarkerNames[index];
            if (dictionary.Items.ContainsKey(marker)) return true;
        }

        foreach (KeyValuePair<string, PdfObject> item in dictionary.Items) {
            if (item.Value is PdfName name) {
                for (int index = 0; index < PdfActiveContentPolicy.MarkerNames.Length; index++) {
                    if (string.Equals(name.Name, PdfActiveContentPolicy.MarkerNames[index], StringComparison.Ordinal)) {
                        return true;
                    }
                }
            }

            if (ContainsActiveContentOutsideWidgets(item.Value, widgetObjectNumbers, visitedReferences)) {
                return true;
            }
        }

        return false;
    }
}
