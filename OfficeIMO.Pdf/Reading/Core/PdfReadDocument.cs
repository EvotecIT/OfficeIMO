namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a parsed PDF document with access to pages, catalog and metadata.
/// Note: MVP reader supports classic xref tables and simple stream parsing sufficient for OfficeIMO.Pdf output.
/// </summary>
public sealed class PdfReadDocument {
    private readonly Dictionary<int, PdfIndirectObject> _objects;
    private readonly string _trailerRaw;
    private readonly PdfReadOptions _options;

    private PdfReadDocument(Dictionary<int, PdfIndirectObject> objects, string trailerRaw, PdfReadOptions? options) {
        _objects = objects; _trailerRaw = trailerRaw; _options = options ?? new PdfReadOptions();
        Pages = CollectPages();
        Metadata = ExtractMetadata();
    }

    /// <summary>All page objects discovered in document order.</summary>
    public IReadOnlyList<PdfReadPage> Pages { get; }

    /// <summary>Document metadata (when present).</summary>
    public PdfMetadata Metadata { get; }

    /// <summary>Loads a PDF from bytes into a typed object model.</summary>
    public static PdfReadDocument Load(byte[] pdf, PdfReadOptions? options = null) {
        var (map, trailer) = PdfSyntax.ParseObjects(pdf);
        return new PdfReadDocument(map, trailer, options);
    }

    /// <summary>Loads a PDF from a file path.</summary>
    public static PdfReadDocument Load(string path, PdfReadOptions? options = null) => Load(File.ReadAllBytes(path), options);

    /// <summary>Extracts full‑document plain text (pages separated by blank lines).</summary>
    public string ExtractText() {
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < Pages.Count; i++) {
            if (i > 0) sb.AppendLine();
            sb.Append(Pages[i].ExtractText());
        }
        return sb.ToString();
    }

    private List<PdfReadPage> CollectPages() {
        // Prefer true page tree traversal when possible (Catalog -> Pages -> Kids ...)
        var result = new List<PdfReadPage>();
        int? catalogId = null;
        foreach (var kv in _objects) {
            if (kv.Value.Value is PdfDictionary d && d.Get<PdfName>("Type")?.Name == "Catalog") { catalogId = kv.Key; break; }
        }
        if (catalogId is int cat && _objects.TryGetValue(cat, out var catObj) && catObj.Value is PdfDictionary catalog) {
            var pagesNode = ResolveDict(catalog.Items.TryGetValue("Pages", out var v) ? v : null);
            if (pagesNode is not null) {
                var kids = ResolveArray(pagesNode.Items.TryGetValue("Kids", out var kidsObj) ? kidsObj : null);
                int kidCount = kids?.Items.Count ?? 0;
                var visited = new HashSet<int>();
                int? target = null;
                var cnt = pagesNode.Get<PdfNumber>("Count");
                if (cnt is not null) {
                    int cc = (int)cnt.Value; if (cc > 0) target = cc;
                }
                TraversePagesNodeDeepLimited(pagesNode, visited, result, target);
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

    private PdfDictionary? ResolveDict(PdfObject? obj) {
        if (obj is null) return null;
        if (obj is PdfDictionary d) return d;
        if (obj is PdfReference r && _objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfDictionary dd) return dd;
        return null;
    }

    private PdfArray? ResolveArray(PdfObject? obj) {
        if (obj is null) return null;
        if (obj is PdfArray a) return a;
        if (obj is PdfReference r && _objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfArray aa) return aa;
        return null;
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
                if (d is not null) TraversePagesNode(d, outList);
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

    private void TraversePagesNodeDeepLimited(PdfDictionary node, HashSet<int> visited, List<PdfReadPage> outList, int? limit) {
        var type = node.Get<PdfName>("Type")?.Name;
        if (type == "Page" || (type is null && IsLikelyPage(node))) {
            int objNum = FindObjectNumberFor(node);
            if (objNum > 0 && visited.Add(objNum)) {
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
            if (t == "Pages" || (t is null && ResolveArray(d.Items.TryGetValue("Kids", out var dKidsObj) ? dKidsObj : null) is not null)) TraversePagesNodeDeepLimited(d, visited, outList, limit);
            else if ((t == "Page" || IsLikelyPage(d) || IsLeafPageByParent(d)) &&
                     (t == "Page" || HasMedia(d) || HasInheritedValue(d, "MediaBox") || HasInheritedValue(d, "CropBox"))) {
                int on = FindObjectNumberFor(d);
                if (on > 0 && visited.Add(on)) {
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
    private int FindObjectNumberFor(PdfDictionary dict) {
        foreach (var kv in _objects) if (ReferenceEquals(kv.Value.Value, dict)) return kv.Key;
        // As a fallback when dictionary was re-parsed separately, match by identity via a simple scan of Page objects
        foreach (var kv in _objects) if (kv.Value.Value is PdfDictionary d && d.Get<PdfName>("Type")?.Name == "Page") return kv.Key;
        return 0;
    }

    private string ToRaw() {
        // Reconstruct raw text for simple metadata extraction without reserialization; ok for small files.
        var sb = new StringBuilder();
        foreach (var kv in _objects.OrderBy(k => k.Key)) {
            sb.Append(kv.Key).Append(" 0 obj\n");
            if (kv.Value.Value is PdfStream s) {
                sb.Append("<< ");
                foreach (var d in s.Dictionary.Items) sb.Append('/').Append(d.Key).Append(' ').Append(' ').Append(' ');
                sb.Append(">>\nstream\n");
                sb.Append(PdfEncoding.Latin1GetString(s.Data)).Append("\nendstream\nendobj\n");
            } else {
                sb.Append("...\nendobj\n");
            }
        }
        sb.Append(_trailerRaw);
        return sb.ToString();
    }

    private PdfMetadata ExtractMetadata() {
        // Trailer has /Info N 0 R when present
        var m = System.Text.RegularExpressions.Regex.Match(_trailerRaw, @"/Info\s+(\d+)\s+0\s+R");
        if (!m.Success) return new PdfMetadata();
        if (!int.TryParse(m.Groups[1].Value, out int infoId)) return new PdfMetadata();
        if (!_objects.TryGetValue(infoId, out var infoObj) || infoObj.Value is not PdfDictionary dict) return new PdfMetadata();
        string? GetStr(string key) => dict.Get<PdfStringObj>(key)?.Value;
        return new PdfMetadata {
            Title = GetStr("Title"),
            Author = GetStr("Author"),
            Subject = GetStr("Subject"),
            Keywords = GetStr("Keywords")
        };
    }
}
