using OfficeIMO.Html.Dom;

namespace OfficeIMO.Chm;

internal sealed partial class ChmNavigationReader {
    private sealed class Item {
        internal string Name = string.Empty;
        internal readonly List<ChmLink> Links = new List<ChmLink>();
        internal readonly List<string> SeeAlso = new List<string>();
        internal readonly List<Item> Children = new List<Item>();
        internal ChmNavigationItem Freeze() => new ChmNavigationItem(Name, Links.AsReadOnly(), SeeAlso.AsReadOnly(), Children.Select(child => child.Freeze()).ToList().AsReadOnly());
    }
    private readonly ChmDocument _book;
    private readonly ChmReadOptions _options;
    private readonly Encoding _encoding;
    private readonly CancellationToken _token;
    private readonly List<(string Title, string Target)> _topics = new List<(string, string)>();
    private int _itemCount;
    internal Dictionary<string, string> TopicTitles { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    internal ChmNavigationReader(ChmDocument book, ChmReadOptions options, Encoding encoding, CancellationToken token) {
        _book = book; _options = options; _encoding = encoding; _token = token;
        ReadTopicTable();
    }

    internal IReadOnlyList<ChmNavigationItem> ReadContents(string? declaredPath) {
        ChmEntry? binary = _book.FindEntry("/#TOCIDX");
        if (binary != null && binary.Length != 0 && _topics.Count != 0) return ReadBinaryContents(binary.GetBytes());
        if (binary != null && binary.Length != 0) _book.Diagnostic("CHM_BINARY_NAVIGATION_TABLES_MISSING", "Compiled contents require the topic, string, and URL tables; an HTML sitemap is used when available.");
        return ReadSitemap(declaredPath, ".hhc");
    }
    internal IReadOnlyList<ChmNavigationItem> ReadIndex(string? declaredPath) {
        ChmEntry? binary = _book.FindEntry("/$WWKeywordLinks/BTree");
        if (binary != null && binary.Length != 0 && _topics.Count != 0) return ReadBinaryIndex(binary.GetBytes());
        if (binary != null && binary.Length != 0) _book.Diagnostic("CHM_BINARY_NAVIGATION_TABLES_MISSING", "Compiled index requires the topic, string, and URL tables; an HTML sitemap is used when available.");
        return ReadSitemap(declaredPath, ".hhk");
    }

    private Item NewItem(int depth) {
        _token.ThrowIfCancellationRequested();
        if (depth > _options.MaxNavigationDepth) throw ChmBinary.Error("NAVIGATION_LIMIT", "The navigation exceeds MaxNavigationDepth.");
        if (_itemCount >= _options.MaxNavigationItems) throw ChmBinary.Error("NAVIGATION_LIMIT", "The navigation exceeds MaxNavigationItems.");
        _itemCount++; return new Item();
    }
    private ChmLink Link(string target, string sourcePath, string? title = null) {
        if (target.Length > _options.MaxPathLength) throw ChmBinary.Error("NAVIGATION_LIMIT", "A navigation reference exceeds MaxPathLength.");
        string? path = ChmPaths.Resolve(target, sourcePath, _options.MaxPathLength);
        string suffix = string.Empty;
        int suffixStart = target.IndexOfAny(new[] { '#', '?' });
        if (suffixStart >= 0) suffix = target.Substring(suffixStart);
        string value = path == null ? target : path + suffix;
        if (path == null) _book.Diagnostic("CHM_EXTERNAL_NAVIGATION", "An external, merged-help, or unsafe navigation target is retained without loading it.", target);
        else if (_book.FindEntry(path) == null) _book.Diagnostic("CHM_TOPIC_MISSING", "A navigation target does not exist in this archive.", value);
        return new ChmLink(value, title);
    }

    private IReadOnlyList<ChmNavigationItem> ReadSitemap(string? declaredPath, string extension) {
        ChmEntry? entry = string.IsNullOrWhiteSpace(declaredPath) ? null : _book.FindEntry(declaredPath!);
        if (entry == null && !string.IsNullOrWhiteSpace(declaredPath))
            _book.Diagnostic("CHM_SITEMAP_MISSING", "The declared navigation sitemap is absent or external.", declaredPath);
        if (entry == null) {
            ChmEntry[] candidates = _book.Resources.Where(resource => resource.Path.EndsWith(extension, StringComparison.OrdinalIgnoreCase)).OrderBy(resource => resource.Path, StringComparer.Ordinal).ToArray();
            if (candidates.Length > 1) _book.Diagnostic("CHM_SITEMAP_AMBIGUOUS", "Several undeclared sitemaps are present; the first ordinal path is used.");
            entry = candidates.FirstOrDefault();
        }
        if (entry == null) return Array.Empty<ChmNavigationItem>();
        HtmlDocument document = _book.ParseSitemap(entry, _token);
        var roots = new List<Item>();
        Process(document.ChildNodes, roots, 1);
        return roots.Select(item => item.Freeze()).ToList().AsReadOnly();

        Item? Process(IEnumerable<HtmlNode> nodes, List<Item> output, int depth) {
            Item? last = null;
            foreach (HtmlNode node in nodes) {
                _token.ThrowIfCancellationRequested();
                if (!(node is HtmlElement element)) continue;
                if (element.LocalName == "object" && string.Equals(element.GetAttribute("type"), "text/sitemap", StringComparison.OrdinalIgnoreCase)) {
                    Item item = NewItem(depth);
                    string? pendingTitle = null;
                    foreach (HtmlElement parameter in element.QuerySelectorAll("param")) {
                        _token.ThrowIfCancellationRequested();
                        string? name = parameter.GetAttribute("name"), value = parameter.GetAttribute("value");
                        if (string.IsNullOrEmpty(value)) continue;
                        if (string.Equals(name, "Name", StringComparison.OrdinalIgnoreCase)) { if (item.Name.Length == 0) item.Name = value!; pendingTitle = value; }
                        else if (string.Equals(name, "Local", StringComparison.OrdinalIgnoreCase)) item.Links.Add(Link(value!, entry.Path, pendingTitle));
                        else if (string.Equals(name, "See Also", StringComparison.OrdinalIgnoreCase)) item.SeeAlso.Add(value!);
                        else if (string.Equals(name, "Merge", StringComparison.OrdinalIgnoreCase)) item.Links.Add(Link(value!, entry.Path));
                    }
                    output.Add(item); last = item;
                } else if (element.LocalName == "ul" && last != null) {
                    Process(element.ChildNodes, last.Children, depth + 1);
                } else {
                    Item? added = Process(element.ChildNodes, output, depth);
                    if (added != null) last = added;
                }
            }
            return last;
        }
    }

    private void ReadTopicTable() {
        ChmEntry? topics = _book.FindEntry("/#TOPICS"), strings = _book.FindEntry("/#STRINGS"), urls = _book.FindEntry("/#URLTBL"), urlStrings = _book.FindEntry("/#URLSTR");
        if (topics == null || strings == null || urls == null || urlStrings == null) return;
        byte[] records = topics.GetBytes(), names = strings.GetBytes(), locations = urls.GetBytes(), references = urlStrings.GetBytes();
        if (records.Length % 16 != 0) throw ChmBinary.Error("TOPIC_TABLE", "The compiled topic table is not a sequence of 16-byte records.");
        if (records.Length / 16 > _options.MaxNavigationItems) throw ChmBinary.Error("NAVIGATION_LIMIT", "The compiled topic table exceeds MaxNavigationItems.");
        for (int position = 0; position < records.Length; position += 16) {
            _token.ThrowIfCancellationRequested();
            uint titleOffset = ChmBinary.U32(records, position + 4);
            string title = titleOffset == uint.MaxValue ? string.Empty : ChmBinary.CString(names, ChmBinary.Index(titleOffset), _encoding, _options.MaxPathLength);
            int location = ChmBinary.Index(ChmBinary.U32(records, position + 8));
            ChmBinary.Range(locations, location, 12);
            int referenceOffset = ChmBinary.Index(ChmBinary.U32(locations, location + 8));
            ChmBinary.Range(references, referenceOffset, 9);
            string target = ChmBinary.CString(references, referenceOffset + 8, _encoding, _options.MaxPathLength);
            _topics.Add((title, target));
            string? path = ChmPaths.Resolve(target, "/", _options.MaxPathLength);
            if (path != null && !TopicTitles.ContainsKey(path) && title.Length > 0) TopicTitles.Add(path, title);
        }
    }
    private (string Title, string Target) Topic(uint index) {
        if (index >= _topics.Count) throw ChmBinary.Error("TOPIC_TABLE", "Compiled navigation refers to an invalid topic index.");
        return _topics[(int)index];
    }
}
