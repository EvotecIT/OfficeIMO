namespace OfficeIMO.Xps;

public sealed partial class XpsDocument {
    internal (int PageIndex, string? Name, string? Uri)? ResolveNavigation(string sourcePart, string uri) {
        if (Uri.TryCreate(uri, UriKind.Absolute, out var absolute) && !uri.StartsWith("/", StringComparison.Ordinal)) {
            return absolute.Scheme == "https" || absolute.Scheme == "http" || absolute.Scheme == "mailto"
                ? (-1, null, uri) : null;
        }
        string[] pieces = uri.Split('#');
        if (pieces.Length > 2) return null;
        string part = pieces[0].Length == 0 ? sourcePart : XpsPackage.Resolve(sourcePart, pieces[0]);
        string? anchor = pieces.Length == 2 ? Uri.UnescapeDataString(pieces[1]) : null;
        int page = LinkTargetPage(part, anchor);
        if (page < 0) return null;
        return (page, anchor != null && _pages[page].HasNamedTarget(anchor) ? anchor : null, null);
    }

    // Stabilize known fixed-page links before sequence edits change positional or document-scoped meaning.
    // Unrecognized extension metadata remains opaque; a removed target remains an explicit unresolved link.
    private Dictionary<XpsPage, XElement> PreserveNavigationTargets(StructureIndex next) {
        var replacements = new Dictionary<XpsPage, XElement>();
        foreach (var page in _pageCache.Values) {
            XElement? copy = null;
            var markup = page.GetMarkup();
            foreach (var attribute in markup.Descendants().Attributes("FixedPage.NavigateUri")) {
                string uri = attribute.Value;
                if (uri.StartsWith("#", StringComparison.Ordinal) || uri.IndexOf(':') >= 0) continue;
                string[] parts = uri.Split('#');
                if (parts.Length > 2) continue;
                string name;
                try { name = XpsPackage.Resolve(page.PartName, parts[0]); } catch (InvalidDataException) { continue; }
                if (!_documentCache.ContainsKey(name) && !name.Equals(_sequence, StringComparison.OrdinalIgnoreCase)) continue;
                string? anchor = parts.Length == 2 ? Uri.UnescapeDataString(parts[1]) : null;
                int previous = LinkTargetPage(name, anchor);
                if (previous < 0 || previous >= _pages.Count) continue;
                XpsPage target = _pages[previous];
                int current = LinkTargetPage(next, name, anchor);
                if (current >= 0 && current < next.Pages.Count && next.Pages[current] == target) continue;
                string fragment = anchor != null && target.HasNamedTarget(anchor) ? "#" + Uri.EscapeDataString(anchor) : "";
                attribute.Value = "/" + target.PartName + fragment;
                copy = markup;
            }
            if (copy != null) replacements.Add(page, copy);
        }
        return replacements;
    }
    internal int LinkTargetPage(string sourcePart, string? anchor) => ResolveNavigationPage(_pages, _documentStarts, _linkTargets, sourcePart, anchor);
    private int LinkTargetPage(StructureIndex index, string part, string? anchor) => ResolveNavigationPage(index.Pages, index.Starts, index.Targets, part, anchor);
    private int ResolveNavigationPage(IReadOnlyList<XpsPage> pages, IReadOnlyDictionary<string, int> starts,
        IReadOnlyDictionary<string, Dictionary<string, int>> targets, string part, string? anchor) {
        if (part.Equals(_sequence, StringComparison.OrdinalIgnoreCase) && int.TryParse(anchor, NumberStyles.None, CultureInfo.InvariantCulture, out int number) && number > 0 && number <= pages.Count) return number - 1;
        if (anchor != null && targets.TryGetValue(part, out var names) && names.TryGetValue(anchor, out int page)) return page;
        if (starts.TryGetValue(part, out int first)) return first < pages.Count ? first : -1;
        if (part.Equals(_sequence, StringComparison.OrdinalIgnoreCase)) return pages.Count > 0 ? 0 : -1;
        for (int i = 0; i < pages.Count; i++) if (pages[i].PartName.Equals(part, StringComparison.OrdinalIgnoreCase)) return i;
        return -1;
    }
}
