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
            System.Console.WriteLine($"CollectPages: Catalog {cat}");
            var pagesNode = ResolveDict(catalog.Items.TryGetValue("Pages", out var v) ? v : null);
            if (pagesNode is not null) {
                var kids = pagesNode.Get<PdfArray>("Kids");
                int kidCount = kids?.Items.Count ?? 0;
                System.Console.WriteLine($"CollectPages: Root /Pages has Kids={kidCount}");
                TraversePagesNode(pagesNode, result);
                System.Console.WriteLine($"CollectPages: Traversal found {result.Count} pages");
            }
        }
        if (result.Count > 0) return result;

        // Fallback: scan all objects and sort by object number
        foreach (var kv in _objects) if (kv.Value.Value is PdfDictionary dict && dict.Get<PdfName>("Type")?.Name == "Page") result.Add(new PdfReadPage(kv.Key, dict, _objects));
        result.Sort((a, b) => a.ObjectNumber.CompareTo(b.ObjectNumber));
        return result;
    }

    private PdfDictionary? ResolveDict(PdfObject? obj) {
        if (obj is null) return null;
        if (obj is PdfDictionary d) return d;
        if (obj is PdfReference r && _objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfDictionary dd) return dd;
        return null;
    }

    private void TraversePagesNode(PdfDictionary node, List<PdfReadPage> outList) {
        var type = node.Get<PdfName>("Type")?.Name;
        if (type == "Page") {
            // Find this node's object number
            int objNum = FindObjectNumberFor(node);
            System.Console.WriteLine($"Traverse: Found Page obj {objNum}");
            outList.Add(new PdfReadPage(objNum, node, _objects));
            return;
        }
        if (type == "Pages") {
            var kids = node.Get<PdfArray>("Kids");
            if (kids is null) return;
            System.Console.WriteLine($"Traverse: /Pages with {kids.Items.Count} kids");
            foreach (var kid in kids.Items) {
                var d = ResolveDict(kid);
                if (d is null) { System.Console.WriteLine("Traverse: Kid unresolved"); continue; }
                if (d is not null) TraversePagesNode(d, outList);
            }
        }
    }

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
