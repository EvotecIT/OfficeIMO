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
                int? target = null;
                var cnt = pagesNode.Get<PdfNumber>("Count");
                if (cnt is not null) {
                    int cc = (int)cnt.Value; if (cc > 0) target = cc;
                }
                TraversePagesNodeDeepLimited(pagesNode, visitedNodes, visitedPages, result, target);
                if (result.Count == 0 && kidCount > 0) {
                    // Build a reachable candidate set from Kids only
                    var reachable = CollectReachableLeafCandidates(pagesNode);
                    foreach (var id in reachable) {
                        if (_objects.TryGetValue(id, out var ind) && ind.Value is PdfDictionary dict) {
                            result.Add(new PdfReadPage(id, dict, _objects));
                            if (target.HasValue && result.Count >= target.Value) break;
                        }
                        if (target.HasValue && result.Count >= target.Value) break;
                    }
                }
            }
        }
        if (result.Count > 0) return result;

        // Fallback: scan all dictionaries; accept leaf candidates whose Parent chain leads to a /Pages node
        foreach (var kv in _objects) {
            if (kv.Value.Value is PdfDictionary dict) {
                if (IsLeafPageByParent(dict)) result.Add(new PdfReadPage(kv.Key, dict, _objects));
            }
        }
        result.Sort((a, b) => a.ObjectNumber.CompareTo(b.ObjectNumber));
        return result;
    }

    private void TraversePagesNode(PdfDictionary node, List<PdfReadPage> outList) {
        var type = node.Get<PdfName>("Type")?.Name;
        if (type == "Page" || (type is null && IsLikelyPage(node))) {
            // Find this node's object number
            int objNum = FindObjectNumberFor(node);
            outList.Add(new PdfReadPage(objNum, node, _objects));
            return;
        }
        var kidsObj = node.Items.TryGetValue("Kids", out var kidsValue) ? kidsValue : null;
        if (type == "Pages" || (type is null && ResolveArray(kidsObj) is not null)) {
            var kids = ResolveArray(kidsObj);
            if (kids is null) return;
            foreach (var kid in kids.Items) {
                var d = ResolveDict(kid);
                if (d is null) { continue; }
                TraversePagesNode(d, outList);
            }
        }
    }

    private bool IsLikelyPage(PdfDictionary d) {
        // Heuristic when /Type is missing: leaf node has Contents, and page data can come from itself or inherited /Pages nodes.
        bool hasContents = d.Items.ContainsKey("Contents");
        bool hasRes = d.Items.ContainsKey("Resources") || HasInheritedValue(d, "Resources");
        bool hasMedia = HasMedia(d) || HasInheritedValue(d, "MediaBox") || HasInheritedValue(d, "CropBox");
        bool hasKids = ResolveArray(d.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null) is not null;
        return !hasKids && hasContents && (hasRes || hasMedia);
    }

    private void TraversePagesNodeDeepLimited(PdfDictionary node, HashSet<PdfDictionary> visitedNodes, HashSet<int> visitedPages, List<PdfReadPage> outList, int? limit) {
        if (!visitedNodes.Add(node)) {
            return;
        }

        var type = node.Get<PdfName>("Type")?.Name;
        if (type == "Page" || (type is null && IsLikelyPage(node))) {
            int objNum = FindObjectNumberFor(node);
            if (objNum > 0 && visitedPages.Add(objNum)) {
                if (type == "Page" || HasMedia(node) || HasInheritedValue(node, "MediaBox") || HasInheritedValue(node, "CropBox")) {
                    outList.Add(new PdfReadPage(objNum, node, _objects));
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
            if (t == "Pages" || (t is null && ResolveArray(d.Items.TryGetValue("Kids", out var dKidsObj) ? dKidsObj : null) is not null)) TraversePagesNodeDeepLimited(d, visitedNodes, visitedPages, outList, limit);
            else if ((t == "Page" || IsLikelyPage(d) || IsLeafPageByParent(d)) &&
                     (t == "Page" || HasMedia(d) || HasInheritedValue(d, "MediaBox") || HasInheritedValue(d, "CropBox"))) {
                int on = FindObjectNumberFor(d);
                if (on > 0 && visitedPages.Add(on)) {
                    outList.Add(new PdfReadPage(on, d, _objects));
                    if (limit.HasValue && outList.Count >= limit.Value) return;
                }
            }
        }
    }

    private HashSet<int> CollectReachableLeafCandidates(PdfDictionary pagesRoot) {
        var set = new HashSet<int>();
        var stack = new Stack<PdfDictionary>();
        stack.Push(pagesRoot);
        int guard = 0;
        while (stack.Count > 0 && guard++ < 10000) {
            var cur = stack.Pop();
            var kids = ResolveArray(cur.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null);
            if (kids is null) continue;
            foreach (var k in kids.Items) {
                var d = ResolveDict(k);
                if (d is null) continue;
                var t = d.Get<PdfName>("Type")?.Name;
                if (t == "Pages" || (t is null && ResolveArray(d.Items.TryGetValue("Kids", out var dKidsObj) ? dKidsObj : null) is not null)) stack.Push(d);
                else if (IsLikelyPage(d) || IsLeafPageByParent(d)) {
                    int on = FindObjectNumberFor(d);
                    if (on > 0) set.Add(on);
                }
            }
        }
        return set;
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
