namespace OfficeIMO.Pdf;

internal sealed partial class PdfPageOptionalContentVisibility {
    private static readonly Dictionary<int, bool> EmptyGroupVisibility = new Dictionary<int, bool>();
    private static readonly HashSet<int> EmptyHiddenObjectNumbers = new HashSet<int>();

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
        int maxExpressionDepth,
        System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        int effectiveMaxExpressionDepth = System.Math.Min(
            maxExpressionDepth,
            PdfReadLimits.DefaultMaxContentNestingDepth);
        if (catalog == null || !catalog.Items.ContainsKey("OCProperties")) {
            return new DocumentState(
                objects,
                EmptyGroupVisibility,
                EmptyHiddenObjectNumbers,
                effectiveMaxExpressionDepth,
                hasUnsupportedViewUsageApplications: false);
        }

        Dictionary<int, bool> groupVisibility = ReadGroupVisibility(
            catalog,
            objects,
            cancellationToken,
            out bool hasUnsupportedViewUsageApplications);
        var hiddenObjectNumbers = new HashSet<int>();
        foreach (KeyValuePair<int, bool> entry in groupVisibility) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!entry.Value) {
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
