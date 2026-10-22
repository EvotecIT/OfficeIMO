namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Indexes final structure membership for one write, retaining both independent parent representations.</summary>
    private sealed class StructureChildIndex {
        // These private layout types have reference equality, matching the
        // existing cross-page ParentElement identity contract.
        private readonly Dictionary<LayoutResult.Page, List<PageStructElement>?[]> _local = new();
        private readonly Dictionary<PageStructElement, List<PageStructElement>> _referenced = new();

        internal StructureChildIndex(IReadOnlyList<LayoutResult.Page> pages) {
            foreach (LayoutResult.Page page in pages) {
                var local = new List<PageStructElement>?[page.StructElements.Count];
                _local[page] = local;
                foreach (PageStructElement child in page.StructElements) {
                    if (child.ParentElementIndex is int index && index >= 0 && index < local.Length) {
                        (local[index] ??= new List<PageStructElement>()).Add(child);
                    }
                    if (child.ParentElement is PageStructElement parent) {
                        if (!_referenced.TryGetValue(parent, out List<PageStructElement>? children)) {
                            children = new List<PageStructElement>();
                            _referenced.Add(parent, children);
                        }
                        children.Add(child);
                    }
                }
            }
        }

        internal List<int> GetChildren(LayoutResult.Page page, PageStructElement element, int elementIndex) {
            List<PageStructElement>? local = _local[page][elementIndex];
            _referenced.TryGetValue(element, out List<PageStructElement>? referenced);
            var children = new List<PageStructElement>((local?.Count ?? 0) + (referenced?.Count ?? 0));
            if (local != null) children.AddRange(local);
            if (referenced != null) children.AddRange(referenced);
            // Local membership comes first. Reference-parent membership follows
            // in page/source order, including entries present in both sets.
            // Sorting owns a fresh list and cannot reorder the stored index.
            return OrderStructureChildren(children);
        }
    }
}
