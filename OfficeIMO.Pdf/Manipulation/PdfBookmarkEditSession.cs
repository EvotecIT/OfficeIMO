namespace OfficeIMO.Pdf;

/// <summary>Editable existing-document bookmark node.</summary>
public sealed class PdfBookmarkNode {
    private readonly List<PdfBookmarkNode> _children = new List<PdfBookmarkNode>();
    internal PdfBookmarkNode(string id, string title, int pageNumber, double? top, bool expanded, PdfOpenActionDestinationMode? mode = null, double? left = null, double? bottom = null, double? right = null, double? zoom = null) { Id = id; Title = title; PageNumber = pageNumber; DestinationTop = top; IsExpanded = expanded; DestinationMode = mode ?? PdfOpenActionDestinationMode.Xyz; DestinationLeft = left; DestinationBottom = bottom; DestinationRight = right; DestinationZoom = zoom; }
    /// <summary>Stable edit-session identifier.</summary>
    public string Id { get; }
    /// <summary>Bookmark title.</summary>
    public string Title { get; internal set; }
    /// <summary>One-based destination page.</summary>
    public int PageNumber { get; internal set; }
    /// <summary>Optional top destination coordinate.</summary>
    public double? DestinationTop { get; internal set; }
    /// <summary>Viewer destination mode.</summary>
    public PdfOpenActionDestinationMode DestinationMode { get; internal set; }
    /// <summary>Optional left destination coordinate.</summary>
    public double? DestinationLeft { get; internal set; }
    /// <summary>Optional bottom destination coordinate.</summary>
    public double? DestinationBottom { get; internal set; }
    /// <summary>Optional right destination coordinate.</summary>
    public double? DestinationRight { get; internal set; }
    /// <summary>Optional XYZ zoom factor.</summary>
    public double? DestinationZoom { get; internal set; }
    /// <summary>Whether children are initially expanded.</summary>
    public bool IsExpanded { get; internal set; }
    /// <summary>Child bookmarks.</summary>
    public IReadOnlyList<PdfBookmarkNode> Children => _children.AsReadOnly();
    internal List<PdfBookmarkNode> MutableChildren => _children;
}

/// <summary>Transactional add/remove/rename/move/nest/retarget/rebuild surface for existing bookmarks.</summary>
public sealed class PdfBookmarkEditSession {
    private readonly List<PdfBookmarkNode> _roots;
    private readonly PdfLogicalDocument _logical;
    private int _nextId;
    internal PdfBookmarkEditSession(PdfLogicalDocument logical) { _logical = logical; _roots = new List<PdfBookmarkNode>(); Import(logical.Outlines, _roots); }
    /// <summary>Current top-level bookmark nodes.</summary>
    public IReadOnlyList<PdfBookmarkNode> Roots => _roots.AsReadOnly();
    /// <summary>Adds a bookmark at the root or below a parent id.</summary>
    public PdfBookmarkNode Add(string title, int pageNumber, string? parentId = null, double? destinationTop = null, bool expanded = true) {
        ValidateTitle(title); ValidatePage(pageNumber); PdfBookmarkNode node = NewNode(title, pageNumber, destinationTop, expanded); GetChildren(parentId).Add(node); return node;
    }
    /// <summary>Removes a bookmark and its descendants.</summary>
    public PdfBookmarkEditSession Remove(string id) { (List<PdfBookmarkNode> siblings, int index) = RequireLocation(id); siblings.RemoveAt(index); return this; }
    /// <summary>Renames a bookmark.</summary>
    public PdfBookmarkEditSession Rename(string id, string title) { ValidateTitle(title); Require(id).Title = title; return this; }
    /// <summary>Moves a bookmark to a root or nested sibling position.</summary>
    public PdfBookmarkEditSession Move(string id, string? parentId = null, int index = -1) {
        (List<PdfBookmarkNode> oldSiblings, int oldIndex) = RequireLocation(id); PdfBookmarkNode node = oldSiblings[oldIndex];
        if (parentId != null && Contains(node, parentId)) throw new InvalidOperationException("A bookmark cannot be moved below itself or one of its descendants.");
        oldSiblings.RemoveAt(oldIndex); List<PdfBookmarkNode> target = GetChildren(parentId); if (index < 0) index = target.Count;
        #pragma warning disable CA1512 // ThrowIfGreaterThan is unavailable on every target framework.
        if (index > target.Count) throw new ArgumentOutOfRangeException(nameof(index));
        #pragma warning restore CA1512
        target.Insert(index, node); return this;
    }
    /// <summary>Retargets a bookmark to a page and optional top coordinate.</summary>
    public PdfBookmarkEditSession Retarget(string id, int pageNumber, double? destinationTop = null) { ValidatePage(pageNumber); PdfBookmarkNode node = Require(id); node.PageNumber = pageNumber; node.DestinationTop = destinationTop; node.DestinationMode = PdfOpenActionDestinationMode.Xyz; node.DestinationLeft = null; node.DestinationBottom = null; node.DestinationRight = null; node.DestinationZoom = null; return this; }
    /// <summary>Retargets a bookmark using an explicit viewer destination mode and coordinates.</summary>
    public PdfBookmarkEditSession Retarget(string id, int pageNumber, PdfOpenActionDestinationMode destinationMode, double? destinationTop = null, double? destinationLeft = null, double? destinationBottom = null, double? destinationRight = null, double? destinationZoom = null) { ValidatePage(pageNumber); PdfBookmarkNode node = Require(id); node.PageNumber = pageNumber; node.DestinationMode = destinationMode; node.DestinationTop = destinationTop; node.DestinationLeft = destinationLeft; node.DestinationBottom = destinationBottom; node.DestinationRight = destinationRight; node.DestinationZoom = destinationZoom; return this; }
    /// <summary>Replaces bookmarks with the source document's inferred heading hierarchy.</summary>
    public PdfBookmarkEditSession RebuildFromHeadings() {
        _roots.Clear(); var stack = new List<PdfBookmarkNode>();
        foreach (PdfLogicalPage page in _logical.Pages) foreach (PdfLogicalHeading heading in page.Headings) {
            int level = Math.Max(1, heading.Level); while (stack.Count >= level) stack.RemoveAt(stack.Count - 1);
            PdfBookmarkNode node = NewNode(heading.Text, page.PageNumber, heading.Line.BaselineY, true);
            (stack.Count == 0 ? _roots : stack[stack.Count - 1].MutableChildren).Add(node); stack.Add(node);
        }
        return this;
    }
    internal IReadOnlyList<PdfBookmarkNode> Snapshot() => _roots;
    private void Import(IReadOnlyList<PdfOutlineItem> source, List<PdfBookmarkNode> target) { foreach (PdfOutlineItem item in source) { if (!item.PageNumber.HasValue) continue; PdfBookmarkNode node = NewNode(item.Title, item.PageNumber.Value, item.DestinationTop, item.IsExpanded, item.DestinationMode, item.DestinationLeft, item.DestinationBottom, item.DestinationRight, item.DestinationZoom); target.Add(node); Import(item.Children, node.MutableChildren); } }
    private PdfBookmarkNode NewNode(string title, int page, double? top, bool expanded, PdfOpenActionDestinationMode? mode = null, double? left = null, double? bottom = null, double? right = null, double? zoom = null) => new PdfBookmarkNode("bookmark-" + (++_nextId).ToString(System.Globalization.CultureInfo.InvariantCulture), title, page, top, expanded, mode, left, bottom, right, zoom);
    private List<PdfBookmarkNode> GetChildren(string? parentId) => parentId == null ? _roots : Require(parentId).MutableChildren;
    private PdfBookmarkNode Require(string id) { Guard.NotNullOrWhiteSpace(id, nameof(id)); return Find(_roots, id) ?? throw new KeyNotFoundException("PDF bookmark was not found: " + id); }
    private (List<PdfBookmarkNode> Siblings, int Index) RequireLocation(string id) { if (TryFindLocation(_roots, id, out List<PdfBookmarkNode>? siblings, out int index)) return (siblings!, index); throw new KeyNotFoundException("PDF bookmark was not found: " + id); }
    private static PdfBookmarkNode? Find(List<PdfBookmarkNode> nodes, string id) { foreach (PdfBookmarkNode node in nodes) { if (node.Id == id) return node; PdfBookmarkNode? child = Find(node.MutableChildren, id); if (child != null) return child; } return null; }
    private static bool TryFindLocation(List<PdfBookmarkNode> nodes, string id, out List<PdfBookmarkNode>? siblings, out int index) { for (int i = 0; i < nodes.Count; i++) { if (nodes[i].Id == id) { siblings = nodes; index = i; return true; } if (TryFindLocation(nodes[i].MutableChildren, id, out siblings, out index)) return true; } siblings = null; index = -1; return false; }
    private static bool Contains(PdfBookmarkNode node, string id) => node.Id == id || node.MutableChildren.Any(child => Contains(child, id));
    private void ValidatePage(int page) { if (page < 1 || page > _logical.Pages.Count) throw new ArgumentOutOfRangeException(nameof(page)); }
    private static void ValidateTitle(string title) => Guard.NotNullOrWhiteSpace(title, nameof(title));
}
