namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private IReadOnlyList<PdfPageLabel> ExtractPageLabels() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue("PageLabels", out var pageLabelsObject) ||
            ResolveObject(pageLabelsObject) is not PdfDictionary tree ||
            tree.Items.ContainsKey("Kids") ||
            !tree.Items.TryGetValue("Nums", out var numsObject) ||
            ResolveObject(numsObject) is not PdfArray nums ||
            nums.Items.Count % 2 != 0) {
            return Array.Empty<PdfPageLabel>();
        }

        var labels = new List<PdfPageLabel>();
        for (int i = 0; i < nums.Items.Count; i += 2) {
            if (ResolveObject(nums.Items[i]) is not PdfNumber pageIndexNumber ||
                !TryGetNonNegativeInteger(pageIndexNumber, out int pageIndex) ||
                ResolveObject(nums.Items[i + 1]) is not PdfDictionary labelDictionary) {
                return Array.Empty<PdfPageLabel>();
            }

            string? style = null;
            if (ResolveObject(labelDictionary.Items.TryGetValue("S", out var styleObject) ? styleObject : null) is PdfName styleName &&
                !string.IsNullOrEmpty(styleName.Name)) {
                style = styleName.Name;
            }

            string? prefix = null;
            if (ResolveObject(labelDictionary.Items.TryGetValue("P", out var prefixObject) ? prefixObject : null) is PdfStringObj prefixText) {
                prefix = prefixText.Value;
            }

            int? startNumber = null;
            if (ResolveObject(labelDictionary.Items.TryGetValue("St", out var startObject) ? startObject : null) is PdfNumber startNumberObject &&
                TryGetPositiveInteger(startNumberObject, out int parsedStartNumber)) {
                startNumber = parsedStartNumber;
            }

            labels.Add(new PdfPageLabel(pageIndex, style, prefix, startNumber));
        }

        labels.Sort((left, right) => left.StartPageIndex.CompareTo(right.StartPageIndex));
        return labels.Count == 0 ? Array.Empty<PdfPageLabel>() : labels.AsReadOnly();
    }
}
