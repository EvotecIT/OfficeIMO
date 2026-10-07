namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Explicit source order rearranges annotated owners only. Ordinary flow and links retain their slots.
    private static List<int> OrderStructureChildren(List<PageStructElement> children) {
        if (children.Count(child => child.LogicalOrder.HasValue) > 1) {
            PageStructElement[] ordered = children.Where(child => child.LogicalOrder.HasValue)
                .OrderBy(child => child.LogicalOrder!.Value).ToArray();
            int next = 0;
            for (int index = 0; index < children.Count; index++) {
                if (children[index].LogicalOrder.HasValue) children[index] = ordered[next++];
            }
        }
        return children.Select(child => child.ObjectId).ToList();
    }
}
