namespace OfficeIMO.Pdf;

internal sealed partial class PdfPageOptionalContentVisibility {
    internal sealed class DocumentState {
        internal DocumentState(
            Dictionary<int, PdfIndirectObject> objects,
            Dictionary<int, bool> groupVisibility,
            HashSet<int> hiddenObjectNumbers,
            int maxExpressionDepth,
            bool hasUnsupportedViewUsageApplications) {
            Objects = objects;
            GroupVisibility = groupVisibility;
            HiddenObjectNumbers = hiddenObjectNumbers;
            MaxExpressionDepth = maxExpressionDepth;
            HasUnsupportedViewUsageApplications = hasUnsupportedViewUsageApplications;
        }

        internal Dictionary<int, PdfIndirectObject> Objects { get; }
        internal Dictionary<int, bool> GroupVisibility { get; }
        internal HashSet<int> HiddenObjectNumbers { get; }
        internal int MaxExpressionDepth { get; }
        internal bool HasUnsupportedViewUsageApplications { get; }
    }

    internal static DocumentState CreateDocumentState(
        PdfDictionary? catalog,
        Dictionary<int, PdfIndirectObject> objects,
        int maxExpressionDepth) {
        int effectiveMaxExpressionDepth = System.Math.Min(
            maxExpressionDepth,
            PdfReadLimits.DefaultMaxContentNestingDepth);
        Dictionary<int, bool> groupVisibility = ReadGroupVisibility(
            catalog,
            objects,
            out bool hasUnsupportedViewUsageApplications);
        var hiddenObjectNumbers = new HashSet<int>();
        foreach (KeyValuePair<int, bool> entry in groupVisibility) {
            if (!entry.Value) {
                hiddenObjectNumbers.Add(entry.Key);
            }
        }

        var visitedReferences = new HashSet<int>();
        foreach (KeyValuePair<int, PdfIndirectObject> entry in objects) {
            if (hiddenObjectNumbers.Contains(entry.Key)) {
                continue;
            }

            visitedReferences.Clear();
            if (IsOptionalContentObjectHidden(
                    entry.Value.Value,
                    groupVisibility,
                    objects,
                    visitedReferences,
                    effectiveMaxExpressionDepth,
                    depth: 0)) {
                hiddenObjectNumbers.Add(entry.Key);
            }
        }

        return new DocumentState(
            objects,
            groupVisibility,
            hiddenObjectNumbers,
            effectiveMaxExpressionDepth,
            hasUnsupportedViewUsageApplications);
    }
}
