namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>
    /// Plans structure leaves whose marked content can belong to their parent without changing what the tree says.
    /// A canvas text Span carries no attributes, so as the only child of a paragraph, heading, label or cell its
    /// content is the parent's own content. A paragraph an adapter generated around loose cell text is likewise the
    /// cell's content. Merged leaves write no object; their object number resolves to the element that owns the content.
    /// </summary>
    private sealed class StructureLeafMerges {
        private readonly Dictionary<PageStructElement, PageStructElement> _mergedInto = new();
        private readonly Dictionary<PageStructElement, PageStructElement> _contentOwner = new();

        private StructureLeafMerges() { }

        internal bool IsMerged(PageStructElement element) => _mergedInto.ContainsKey(element);

        /// <summary>The merged leaf whose marked content this element writes, or null when it keeps its own.</summary>
        internal PageStructElement? ContentOwner(PageStructElement element) =>
            _contentOwner.TryGetValue(element, out PageStructElement? leaf) ? leaf : null;

        internal void AliasMergedObjectIds() {
            foreach (PageStructElement leaf in _mergedInto.Keys) {
                PageStructElement owner = leaf;
                while (_mergedInto.TryGetValue(owner, out PageStructElement? parent)) owner = parent;
                leaf.ObjectId = owner.ObjectId;
            }
        }

        internal static StructureLeafMerges? Plan(IReadOnlyList<LayoutResult.Page> pages) {
            bool hasCandidate = false;
            foreach (LayoutResult.Page page in pages) {
                foreach (PageStructElement element in page.StructElements) {
                    if (element.IsPlainCanvasTextSpan || element.IsAnonymousBlock) {
                        hasCandidate = true;
                        break;
                    }
                }
                if (hasCandidate) break;
            }
            if (!hasCandidate) return null;

            var children = new Dictionary<PageStructElement, List<PageStructElement>>();
            var parents = new Dictionary<PageStructElement, PageStructElement?>();
            var pageOf = new Dictionary<PageStructElement, LayoutResult.Page>();
            var multiPage = new HashSet<PageStructElement>();
            foreach (LayoutResult.Page page in pages) {
                foreach (PageStructElement child in page.StructElements) {
                    if (pageOf.TryGetValue(child, out _)) multiPage.Add(child);
                    else pageOf[child] = page;
                    PageStructElement? local = child.ParentElementIndex is int index && index >= 0 && index < page.StructElements.Count
                        ? page.StructElements[index] : null;
                    AddChild(children, local, child);
                    AddChild(children, child.ParentElement, child);
                    // A leaf with two different parents is left as written.
                    parents[child] = local != null && child.ParentElement != null && !ReferenceEquals(local, child.ParentElement)
                        ? null : local ?? child.ParentElement;
                }
            }

            var plan = new StructureLeafMerges();
            foreach (LayoutResult.Page page in pages) {
                foreach (PageStructElement leaf in page.StructElements) {
                    if (!leaf.IsPlainCanvasTextSpan || leaf.StructureType != "Span" || !leaf.MarkedContentId.HasValue
                        || children.ContainsKey(leaf)) continue;
                    plan.TryMerge(leaf, leaf, parents, children, pageOf, multiPage, IsTextOwner);
                }
            }
            foreach (LayoutResult.Page page in pages) {
                foreach (PageStructElement paragraph in page.StructElements) {
                    if (!paragraph.IsAnonymousBlock || paragraph.StructureType != "P" || plan.IsMerged(paragraph)) continue;
                    PageStructElement? content = paragraph.MarkedContentId.HasValue && !children.ContainsKey(paragraph)
                        ? paragraph : plan.ContentOwner(paragraph);
                    if (content == null) continue;
                    plan.TryMerge(paragraph, content, parents, children, pageOf, multiPage, IsCellOwner);
                }
            }
            return plan._mergedInto.Count == 0 ? null : plan;
        }

        private void TryMerge(PageStructElement leaf, PageStructElement content,
            Dictionary<PageStructElement, PageStructElement?> parents,
            Dictionary<PageStructElement, List<PageStructElement>> children,
            Dictionary<PageStructElement, LayoutResult.Page> pageOf, HashSet<PageStructElement> multiPage,
            Func<string, bool> acceptsOwner) {
            if (!HasNoOwnAttributes(leaf) || multiPage.Contains(leaf)
                || !parents.TryGetValue(leaf, out PageStructElement? owner) || owner == null
                || multiPage.Contains(owner) || !acceptsOwner(owner.StructureType) || owner.MarkedContentId.HasValue
                || !HasNoOwnAttributes(owner) || _mergedInto.ContainsKey(owner) || _contentOwner.ContainsKey(owner)
                || !children.TryGetValue(owner, out List<PageStructElement>? siblings) || siblings.Count != 1
                || !pageOf.TryGetValue(owner, out LayoutResult.Page? ownerPage)
                || !pageOf.TryGetValue(leaf, out LayoutResult.Page? leafPage) || !ReferenceEquals(ownerPage, leafPage)
                || !string.Equals(leaf.Language, owner.Language, StringComparison.Ordinal)) return;
            _mergedInto.Add(leaf, owner);
            _contentOwner.Add(owner, content);
        }

        private static bool HasNoOwnAttributes(PageStructElement element) =>
            element.AnnotationObjectId == null && element.AdditionalAnnotationObjectIds == null
            && element.AlternativeText.Length == 0 && element.AssociatedFiles.Count == 0
            && element.ListNumbering == null && !element.IsLinkLeadingWhitespace && !element.SpansPages;

        private static bool IsTextOwner(string structureType) => structureType is "P" or "H" or "H1" or "H2" or "H3"
            or "H4" or "H5" or "H6" or "Lbl" or "LBody" or "Caption" or "TD" or "TH";

        private static bool IsCellOwner(string structureType) => structureType is "TD" or "TH";

        private static void AddChild(Dictionary<PageStructElement, List<PageStructElement>> children,
            PageStructElement? parent, PageStructElement child) {
            if (parent == null) return;
            if (!children.TryGetValue(parent, out List<PageStructElement>? list)) {
                list = new List<PageStructElement>();
                children.Add(parent, list);
            }
            // Both parent links of one child are recorded consecutively.
            if (list.Count == 0 || !ReferenceEquals(list[list.Count - 1], child)) list.Add(child);
        }
    }
}
