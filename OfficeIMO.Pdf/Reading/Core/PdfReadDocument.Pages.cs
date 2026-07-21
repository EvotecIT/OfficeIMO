namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private List<PdfReadPage> CollectPages() {
        // Prefer true page tree traversal when possible (Catalog -> Pages -> Kids ...)
        var result = new List<PdfReadPage>();
        PdfDictionary? catalog = FindCatalog();
        if (catalog is not null) {
            var pagesNode = ResolveDict(catalog.Items.TryGetValue("Pages", out var v) ? v : null);
            if (pagesNode is not null) {
                var kids = ResolveArray(pagesNode.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null);
                int kidCount = kids?.Items.Count ?? 0;
                var visitedNodes = new HashSet<PdfDictionary>();
                var visitedPages = new HashSet<int>();
                TraversePagesNodeDeepLimited(pagesNode, visitedNodes, visitedPages, result, limit: null, depth: 1);
                if (result.Count == 0 && kidCount > 0) {
                    // Build a reachable candidate set from Kids only
                    var reachable = CollectReachableLeafCandidates(pagesNode);
                    foreach (var id in reachable) {
                        if (_objects.TryGetValue(id, out var ind) && ind.Value is PdfDictionary dict) {
                            AddPageWithinBudget(result, CreateReadPage(id, dict));
                        }
                    }
                }
            }
        }
        if (result.Count > 0) return result;

        // Fallback: scan all dictionaries; accept leaf candidates whose Parent chain leads to a /Pages node
        foreach (var kv in _objects) {
            if (kv.Value.Value is PdfDictionary dict) {
                if (IsLeafPageByParent(dict)) AddPageWithinBudget(result, CreateReadPage(kv.Key, dict));
            }
        }
        result.Sort((a, b) => a.ObjectNumber.CompareTo(b.ObjectNumber));
        return result;
    }

    private bool IsLikelyPage(PdfDictionary d) {
        // Heuristic when /Type is missing: leaf node has Contents, and page data can come from itself or inherited /Pages nodes.
        bool hasContents = d.Items.ContainsKey("Contents");
        bool hasRes = d.Items.ContainsKey("Resources") || HasInheritedValue(d, "Resources");
        bool hasMedia = HasMedia(d) || HasInheritedValue(d, "MediaBox") || HasInheritedValue(d, "CropBox");
        bool hasKids = ResolveArray(d.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null) is not null;
        return !hasKids && hasContents && (hasRes || hasMedia);
    }

    private void TraversePagesNodeDeepLimited(PdfDictionary node, HashSet<PdfDictionary> visitedNodes, HashSet<int> visitedPages, List<PdfReadPage> outList, int? limit, int depth) {
        if (depth > _options.Limits.MaxPageTreeDepth) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.PageTreeDepth, _options.Limits.MaxPageTreeDepth, depth);
        }

        if (!visitedNodes.Add(node)) {
            return;
        }
        if (visitedNodes.Count > _options.Limits.MaxPageTreeNodes) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.PageTreeNodes, _options.Limits.MaxPageTreeNodes, visitedNodes.Count);
        }

        var type = node.Get<PdfName>("Type")?.Name;
        if (type == "Page" || (type is null && IsLikelyPage(node))) {
            int objNum = FindObjectNumberFor(node);
            if (objNum > 0 && visitedPages.Add(objNum)) {
                if (type == "Page" || HasMedia(node) || HasInheritedValue(node, "MediaBox") || HasInheritedValue(node, "CropBox")) {
                    AddPageWithinBudget(outList, CreateReadPage(objNum, node));
                }
            }
            return;
        }
        var kids = ResolveArray(node.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null);
        if (kids is null) return;
        foreach (var kid in kids.Items) {
            if (limit.HasValue && outList.Count >= limit.Value) return;
            var d = ResolveDict(kid);
            if (d is null) { continue; }
            var t = d.Get<PdfName>("Type")?.Name;
            if (t == "Pages" || (t is null && ResolveArray(d.Items.TryGetValue("Kids", out var dKidsObj) ? dKidsObj : null) is not null)) TraversePagesNodeDeepLimited(d, visitedNodes, visitedPages, outList, limit, depth + 1);
            else if ((t == "Page" || IsLikelyPage(d) || IsLeafPageByParent(d)) &&
                     (t == "Page" || HasMedia(d) || HasInheritedValue(d, "MediaBox") || HasInheritedValue(d, "CropBox"))) {
                int on = FindObjectNumberFor(d);
                if (on > 0 && visitedPages.Add(on)) {
                    AddPageWithinBudget(outList, CreateReadPage(on, d));
                    if (limit.HasValue && outList.Count >= limit.Value) return;
                }
            }
        }
    }

    private PdfReadPage CreateReadPage(int objectNumber, PdfDictionary pageDictionary) =>
        new PdfReadPage(objectNumber, pageDictionary, _objects, _options.Limits, DemandTextExtraction, DemandContentExtraction);

    private HashSet<int> CollectReachableLeafCandidates(PdfDictionary pagesRoot) {
        var set = new HashSet<int>();
        var visited = new HashSet<PdfDictionary>();
        var stack = new Stack<(PdfDictionary Node, int Depth)>();
        stack.Push((pagesRoot, 1));
        while (stack.Count > 0) {
            (PdfDictionary cur, int depth) = stack.Pop();
            if (depth > _options.Limits.MaxPageTreeDepth) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.PageTreeDepth, _options.Limits.MaxPageTreeDepth, depth);
            }

            if (!visited.Add(cur)) {
                continue;
            }

            if (visited.Count > _options.Limits.MaxPageTreeNodes) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.PageTreeNodes, _options.Limits.MaxPageTreeNodes, visited.Count);
            }

            var kids = ResolveArray(cur.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null);
            if (kids is null) continue;
            foreach (var k in kids.Items) {
                var d = ResolveDict(k);
                if (d is null) continue;
                var t = d.Get<PdfName>("Type")?.Name;
                if (t == "Pages" || (t is null && ResolveArray(d.Items.TryGetValue("Kids", out var dKidsObj) ? dKidsObj : null) is not null)) stack.Push((d, depth + 1));
                else if (IsLikelyPage(d) || IsLeafPageByParent(d)) {
                    int on = FindObjectNumberFor(d);
                    if (on > 0 && set.Add(on) && set.Count > _options.Limits.MaxPages) {
                        throw PdfReadLimitException.Create(PdfReadLimitKind.Pages, _options.Limits.MaxPages, set.Count);
                    }
                }
            }
        }
        return set;
    }

    private void AddPageWithinBudget(List<PdfReadPage> pages, PdfReadPage page) {
        if (pages.Count >= _options.Limits.MaxPages) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.Pages, _options.Limits.MaxPages, pages.Count + 1L);
        }

        pages.Add(page);
    }
    private bool IsLeafPageByParent(PdfDictionary d) {
        if (!IsLikelyPage(d)) return false;
        // Follow Parent chain up until /Pages or no parent
        PdfDictionary? current = d;
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (!current.Items.TryGetValue("Parent", out var p)) break;
            var parent = ResolveDict(p);
            if (parent is null) break;
            var type = parent.Get<PdfName>("Type")?.Name;
            if (type == "Pages") return true;
            current = parent;
        }
        return false;
    }

    private bool HasInheritedValue(PdfDictionary start, string key) {
        PdfDictionary? current = start;
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (current.Items.ContainsKey(key)) {
                return true;
            }

            if (!current.Items.TryGetValue("Parent", out var parentObj)) {
                break;
            }

            var parent = ResolveDict(parentObj);
            if (parent is null) {
                break;
            }

            current = parent;
        }

        return false;
    }

    private static bool HasMedia(PdfDictionary d) => d.Items.ContainsKey("MediaBox") || d.Items.ContainsKey("CropBox");
}
