namespace OfficeIMO.Pdf;

/// <summary>One bookmark destination validation issue.</summary>
public sealed class PdfBookmarkValidationIssue {
    internal PdfBookmarkValidationIssue(string title, string code, string message) { Title = title; Code = code; Message = message; }
    /// <summary>Bookmark title.</summary>
    public string Title { get; }
    /// <summary>Stable issue code.</summary>
    public string Code { get; }
    /// <summary>Actionable explanation.</summary>
    public string Message { get; }
}

/// <summary>Existing-document bookmark edit result.</summary>
public sealed class PdfBookmarkEditResult {
    private readonly byte[] _pdf;
    internal PdfBookmarkEditResult(byte[] pdf, PdfMutationPlan plan, IReadOnlyList<PdfOutlineItem> outlines) { _pdf = (byte[])pdf.Clone(); MutationPlan = plan; Outlines = outlines; }
    /// <summary>Shared full-rewrite mutation plan.</summary>
    public PdfMutationPlan MutationPlan { get; }
    /// <summary>Bookmarks read back from the saved artifact.</summary>
    public IReadOnlyList<PdfOutlineItem> Outlines { get; }
    /// <summary>Returns edited PDF bytes.</summary>
    public byte[] ToBytes() => (byte[])_pdf.Clone();
    /// <summary>Opens the edited artifact.</summary>
    public PdfDocument ToDocument() => PdfDocument.Open(_pdf);
}

/// <summary>Adds, removes, renames, moves, nests, retargets, and rebuilds existing-document bookmarks.</summary>
internal static class PdfBookmarkEditor {
    /// <summary>Reports bookmarks whose destinations cannot be resolved to a current page.</summary>
    public static IReadOnlyList<PdfBookmarkValidationIssue> Validate(byte[] pdf, PdfReadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf)); PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions); var issues = new List<PdfBookmarkValidationIssue>(); CollectIssues(document.Outlines, document.Pages.Count, issues); return issues.AsReadOnly();
    }

    /// <summary>Applies a transactional bookmark edit and validates exact title, hierarchy, and target readback.</summary>
    public static PdfBookmarkEditResult Edit(byte[] pdf, Action<PdfBookmarkEditSession> edit, PdfReadOptions? readOptions = null) {
        return EditCore(pdf, edit, readOptions, allowBrokenSourceDestinations: false);
    }

    internal static PdfBookmarkEditResult EditAllowingBrokenSourceDestinations(byte[] pdf, Action<PdfBookmarkEditSession> edit) {
        return EditCore(pdf, edit, readOptions: null, allowBrokenSourceDestinations: true);
    }

    private static PdfBookmarkEditResult EditCore(byte[] pdf, Action<PdfBookmarkEditSession> edit, PdfReadOptions? readOptions, bool allowBrokenSourceDestinations) {
        Guard.NotNull(pdf, nameof(pdf)); Guard.NotNull(edit, nameof(edit));
        PdfMutationPlan plan = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyCatalog, readOptions);
        PdfReadDocument read = PdfReadDocument.Open(pdf, readOptions); PdfLogicalDocument logical = PdfLogicalDocument.From(read);
        IReadOnlyList<PdfBookmarkValidationIssue> sourceIssues = Validate(pdf, readOptions);
        if (!allowBrokenSourceDestinations && sourceIssues.Count > 0) throw new InvalidOperationException("PDF bookmark editing requires broken destinations to be repaired or removed first: " + string.Join(" ", sourceIssues.Select(static issue => issue.Message)));
        var session = new PdfBookmarkEditSession(logical); edit(session); IReadOnlyList<PdfBookmarkNode> target = session.Snapshot();
        byte[] output = PdfDocumentObjectGraphRewriter.Rewrite(pdf, readOptions, null, (objects, security) => { RewriteOutlines(objects, security, read, target); return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value) ? security.InfoObjectNumber : null; });
        IReadOnlyList<PdfOutlineItem> actual = PdfReadDocument.Open(output).Outlines;
        if (!Matches(target, actual)) throw new InvalidOperationException("PDF bookmark post-save validation failed; the artifact was not returned.");
        return new PdfBookmarkEditResult(output, plan, actual);
    }

    private static void RewriteOutlines(Dictionary<int, PdfIndirectObject> objects, PdfDocumentSecurityInfo security, PdfReadDocument document, IReadOnlyList<PdfBookmarkNode> roots) {
        if (!security.RootObjectNumber.HasValue || !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? root) || root.Value is not PdfDictionary catalog) throw new InvalidOperationException("PDF catalog is not readable.");
        catalog.Items.Remove("Outlines"); if (roots.Count == 0) return;
        int next = objects.Keys.Max() + 1; int rootNumber = next++; var ids = new Dictionary<PdfBookmarkNode, int>(); Assign(roots, ids, ref next);
        var outlineRoot = new PdfDictionary(); outlineRoot.Items["Type"] = new PdfName("Outlines"); outlineRoot.Items["First"] = new PdfReference(ids[roots[0]], 0); outlineRoot.Items["Last"] = new PdfReference(ids[roots[roots.Count - 1]], 0); outlineRoot.Items["Count"] = new PdfNumber(roots.Sum(VisibleCount)); objects[rootNumber] = new PdfIndirectObject(rootNumber, 0, outlineRoot);
        WriteSiblings(objects, document, roots, rootNumber, ids); catalog.Items["Outlines"] = new PdfReference(rootNumber, 0);
    }

    private static void Assign(IReadOnlyList<PdfBookmarkNode> nodes, Dictionary<PdfBookmarkNode, int> ids, ref int next) { foreach (PdfBookmarkNode node in nodes) { ids[node] = next++; Assign(node.Children, ids, ref next); } }
    private static void WriteSiblings(Dictionary<int, PdfIndirectObject> objects, PdfReadDocument document, IReadOnlyList<PdfBookmarkNode> nodes, int parent, Dictionary<PdfBookmarkNode, int> ids) {
        for (int i = 0; i < nodes.Count; i++) { PdfBookmarkNode node = nodes[i]; var dictionary = new PdfDictionary(); dictionary.Items["Title"] = new PdfStringObj(node.Title, true); dictionary.Items["Parent"] = new PdfReference(parent, 0); if (i > 0) dictionary.Items["Prev"] = new PdfReference(ids[nodes[i - 1]], 0); if (i + 1 < nodes.Count) dictionary.Items["Next"] = new PdfReference(ids[nodes[i + 1]], 0);
            int pageObject = document.Pages[node.PageNumber - 1].ObjectNumber; dictionary.Items["Dest"] = BuildDestination(node, pageObject);
            if (node.Children.Count > 0) { dictionary.Items["First"] = new PdfReference(ids[node.Children[0]], 0); dictionary.Items["Last"] = new PdfReference(ids[node.Children[node.Children.Count - 1]], 0); int descendants = node.Children.Sum(VisibleCount); dictionary.Items["Count"] = new PdfNumber(node.IsExpanded ? descendants : -descendants); }
            objects[ids[node]] = new PdfIndirectObject(ids[node], 0, dictionary); WriteSiblings(objects, document, node.Children, ids[node], ids); }
    }
    private static PdfArray BuildDestination(PdfBookmarkNode node, int pageObject) {
        var destination = new PdfArray(); destination.Items.Add(new PdfReference(pageObject, 0));
        switch (node.DestinationMode) {
            case PdfOpenActionDestinationMode.Xyz: destination.Items.Add(new PdfName("XYZ")); AddNumber(destination, node.DestinationLeft); AddNumber(destination, node.DestinationTop); AddNumber(destination, node.DestinationZoom); break;
            case PdfOpenActionDestinationMode.Fit: destination.Items.Add(new PdfName("Fit")); break;
            case PdfOpenActionDestinationMode.FitHorizontal: destination.Items.Add(new PdfName("FitH")); AddNumber(destination, node.DestinationTop); break;
            case PdfOpenActionDestinationMode.FitVertical: destination.Items.Add(new PdfName("FitV")); AddNumber(destination, node.DestinationLeft); break;
            case PdfOpenActionDestinationMode.FitRectangle: destination.Items.Add(new PdfName("FitR")); AddNumber(destination, node.DestinationLeft); AddNumber(destination, node.DestinationBottom); AddNumber(destination, node.DestinationRight); AddNumber(destination, node.DestinationTop); break;
            case PdfOpenActionDestinationMode.FitBoundingBox: destination.Items.Add(new PdfName("FitB")); break;
            case PdfOpenActionDestinationMode.FitBoundingBoxHorizontal: destination.Items.Add(new PdfName("FitBH")); AddNumber(destination, node.DestinationTop); break;
            case PdfOpenActionDestinationMode.FitBoundingBoxVertical: destination.Items.Add(new PdfName("FitBV")); AddNumber(destination, node.DestinationLeft); break;
            default: throw new ArgumentOutOfRangeException(nameof(node), node.DestinationMode, "PDF bookmark destination mode is not supported.");
        }
        return destination;
    }
    private static void AddNumber(PdfArray destination, double? value) => destination.Items.Add(value.HasValue ? new PdfNumber(value.Value) : PdfNull.Instance);
    private static int VisibleCount(PdfBookmarkNode node) => 1 + (node.IsExpanded ? node.Children.Sum(VisibleCount) : 0);
    private static bool Matches(IReadOnlyList<PdfBookmarkNode> expected, IReadOnlyList<PdfOutlineItem> actual) { if (expected.Count != actual.Count) return false; for (int i = 0; i < expected.Count; i++) if (expected[i].Title != actual[i].Title || expected[i].PageNumber != actual[i].PageNumber || expected[i].DestinationMode != actual[i].DestinationMode || expected[i].DestinationTop != actual[i].DestinationTop || expected[i].DestinationLeft != actual[i].DestinationLeft || expected[i].DestinationBottom != actual[i].DestinationBottom || expected[i].DestinationRight != actual[i].DestinationRight || expected[i].DestinationZoom != actual[i].DestinationZoom || expected[i].Children.Count != actual[i].Children.Count || !Matches(expected[i].Children, actual[i].Children)) return false; return true; }
    private static void CollectIssues(IReadOnlyList<PdfOutlineItem> outlines, int pageCount, List<PdfBookmarkValidationIssue> issues) { foreach (PdfOutlineItem outline in outlines) { if (!outline.PageNumber.HasValue || outline.PageNumber < 1 || outline.PageNumber > pageCount) issues.Add(new PdfBookmarkValidationIssue(outline.Title, "BrokenDestination", "Bookmark '" + outline.Title + "' does not resolve to a current page.")); CollectIssues(outline.Children, pageCount, issues); } }
}
