namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static readonly XNamespace Ncx = "http://www.daisy.org/z3986/2005/ncx/";

    /// <summary>Replaces TOC entries and optionally page-list/landmark entries while retaining unrelated navigation XML.</summary>
    public void SetNavigation(IEnumerable<EpubNavigationEntry> tableOfContents,
        IEnumerable<EpubNavigationEntry>? pageList = null, IEnumerable<EpubNavigationEntry>? landmarks = null) {
        if (tableOfContents == null) throw new ArgumentNullException(nameof(tableOfContents));
        EpubNavigationEntry[] toc = tableOfContents.ToArray();
        EpubNavigationEntry[]? pages = pageList?.ToArray();
        EpubNavigationEntry[]? guide = landmarks?.ToArray();
        string path = NavigationPath();
        XDocument navigation = ParseXml(_entries[path], 64L * 1024 * 1024);
        XElement? newGuide = null;
        if (PackageVersion == "3.0") {
            XElement body = navigation.Root?.Element(Html + "body") ?? throw new InvalidDataException("Navigation has no XHTML body.");
            SetHtmlNavigation(body, "toc", "Contents", toc, path);
            if (pages != null) SetHtmlNavigation(body, "page-list", "Pages", pages, path);
            if (guide != null) SetHtmlNavigation(body, "landmarks", "Landmarks", guide, path);
        } else {
            XElement root = navigation.Root ?? throw new InvalidDataException("NCX has no root.");
            if (root.Name != Ncx + "ncx") throw new InvalidDataException("Expected an NCX document.");
            XElement map = root.Element(Ncx + "navMap") ?? throw new InvalidDataException("NCX has no navMap.");
            int order = 0;
            map.ReplaceNodes(BuildNcxNodes(toc, path, 0, ref order));
            XElement? depth = root.Element(Ncx + "head")?.Elements(Ncx + "meta").FirstOrDefault(item => (string?)item.Attribute("name") == "dtb:depth");
            depth?.SetAttributeValue("content", NavigationDepth(toc).ToString(System.Globalization.CultureInfo.InvariantCulture));
            if (pages != null) {
                XElement? old = root.Element(Ncx + "pageList"); old?.Remove();
                if (pages.Length != 0) root.Add(new XElement(Ncx + "pageList", new XElement(Ncx + "navLabel", new XElement(Ncx + "text", "Pages")),
                    pages.Select((page, index) => new XElement(Ncx + "pageTarget", new XAttribute("id", "page-" + (index + 1)),
                        new XAttribute("playOrder", order + index + 1),
                        new XAttribute("type", page.SemanticType ?? "normal"), new XAttribute("value", index + 1),
                        new XElement(Ncx + "navLabel", new XElement(Ncx + "text", page.Label)),
                        new XElement(Ncx + "content", new XAttribute("src", NavigationHref(path, page)))))));
                foreach (string name in new[] { "dtb:totalPageCount", "dtb:maxPageNumber" }) {
                    XElement? count = root.Element(Ncx + "head")?.Elements(Ncx + "meta").FirstOrDefault(meta => (string?)meta.Attribute("name") == name);
                    count?.SetAttributeValue("content", pages.Length.ToString(System.Globalization.CultureInfo.InvariantCulture));
                }
            }
            NormalizeNcxPlayOrder(navigation, path);
            if (guide != null) {
                if (guide.Length != 0) newGuide = new XElement(Opf + "guide", guide.Select(item => new XElement(Opf + "reference",
                    new XAttribute("type", item.SemanticType ?? "text"), new XAttribute("title", item.Label),
                    new XAttribute("href", NavigationHref(PackagePath, item)))));
            }
        }
        ReplaceResourcePayload(path, SerializeXml(navigation));
        if (PackageVersion == "2.0" && guide != null) {
            Root.Element(Opf + "guide")?.Remove();
            if (newGuide != null) Root.Add(newGuide);
        }
    }

    private void InitializeNavigation() {
        if (PackageVersion == "3.0") {
            var navigation = new XDocument(new XElement(Html + "html", new XAttribute(XNamespace.Xml + "lang", Language),
                new XAttribute(XNamespace.Xmlns + "epub", Ops.NamespaceName),
                new XElement(Html + "head", new XElement(Html + "title", "Contents")),
                new XElement(Html + "body", new XElement(Html + "nav", new XAttribute(Ops + "type", "toc"),
                    new XElement(Html + "h1", "Contents"), new XElement(Html + "ol")))));
            AddResource("navigation", "EPUB/nav.xhtml", "application/xhtml+xml", SerializeXml(navigation), "nav");
        } else {
            var navigation = new XDocument(new XElement(Ncx + "ncx", new XAttribute("version", "2005-1"),
                new XElement(Ncx + "head", new XElement(Ncx + "meta", new XAttribute("name", "dtb:uid"), new XAttribute("content", Identifier)),
                    new XElement(Ncx + "meta", new XAttribute("name", "dtb:depth"), new XAttribute("content", "1")),
                    new XElement(Ncx + "meta", new XAttribute("name", "dtb:totalPageCount"), new XAttribute("content", "0")),
                    new XElement(Ncx + "meta", new XAttribute("name", "dtb:maxPageNumber"), new XAttribute("content", "0"))),
                new XElement(Ncx + "docTitle", new XElement(Ncx + "text", Title)), new XElement(Ncx + "navMap")));
            AddResource("navigation", "EPUB/toc.ncx", "application/x-dtbncx+xml", SerializeXml(navigation));
            RequireSection("spine").SetAttributeValue("toc", "navigation");
        }
    }
    private string NavigationPath() {
        EpubManifestItem? item = PackageVersion == "3.0" ? Manifest.SingleOrDefault(resource => HasToken(resource.Properties, "nav")) :
            Manifest.SingleOrDefault(resource => resource.Id == (string?)RequireSection("spine").Attribute("toc"));
        if (item == null) throw new InvalidDataException("Package has no declared navigation resource.");
        string path = RequireLocalPath(item);
        if (!_entries.ContainsKey(path)) throw new InvalidDataException("Navigation resource is missing.");
        return path;
    }
    private byte[] PrepareAppendedNavigation(EpubNavigationEntry entry, string pendingPath) {
        string path = NavigationPath();
        XDocument navigation = ParseXml(_entries[path], 64L * 1024 * 1024);
        if (PackageVersion == "3.0") {
            XElement nav = navigation.Descendants(Html + "nav").Single(element => HasToken((string?)element.Attribute(Ops + "type"), "toc"));
            XElement list = nav.Element(Html + "ol") ?? throw new InvalidDataException("TOC list is missing.");
            list.Add(BuildHtmlNodes(new[] { entry }, path, 0, pendingPath));
        } else {
            XElement map = navigation.Root?.Element(Ncx + "navMap") ?? throw new InvalidDataException("NCX navMap is missing.");
            int order = map.Descendants(Ncx + "navPoint").Count();
            map.Add(BuildNcxNodes(new[] { entry }, path, 0, ref order, pendingPath));
            NormalizeNcxPlayOrder(navigation, path);
        }
        return SerializeXml(navigation);
    }
    private void SetHtmlNavigation(XElement body, string type, string heading, EpubNavigationEntry[] nodes, string path) {
        XElement? nav = body.Descendants(Html + "nav").FirstOrDefault(element => HasToken((string?)element.Attribute(Ops + "type"), type));
        if (nav == null) {
            nav = new XElement(Html + "nav", new XAttribute(Ops + "type", type), new XElement(Html + "h1", heading));
            body.Add(nav);
        }
        XElement list = new XElement(Html + "ol", BuildHtmlNodes(nodes, path, 0));
        XElement? old = nav.Element(Html + "ol");
        if (old != null) old.ReplaceWith(list); else nav.Add(list);
    }
    private IEnumerable<XElement> BuildHtmlNodes(IEnumerable<EpubNavigationEntry> nodes, string path, int depth, string? pendingPath = null) {
        if (depth > 64) throw new InvalidDataException("Navigation depth exceeds 64.");
        foreach (EpubNavigationEntry node in nodes) {
            var anchor = new XElement(Html + "a", new XAttribute("href", NavigationHref(path, node, pendingPath)), node.Label);
            anchor.SetAttributeValue(Ops + "type", node.SemanticType);
            yield return new XElement(Html + "li", anchor, node.Children.Count == 0 ? null :
                new XElement(Html + "ol", BuildHtmlNodes(node.Children, path, depth + 1, pendingPath)));
        }
    }
    private IEnumerable<XElement> BuildNcxNodes(IEnumerable<EpubNavigationEntry> nodes, string path, int depth, ref int order, string? pendingPath = null) {
        if (depth > 64) throw new InvalidDataException("Navigation depth exceeds 64.");
        var result = new List<XElement>();
        foreach (EpubNavigationEntry node in nodes) {
            int current = ++order;
            result.Add(new XElement(Ncx + "navPoint", new XAttribute("id", "nav-" + current), new XAttribute("playOrder", current),
                new XElement(Ncx + "navLabel", new XElement(Ncx + "text", node.Label)),
                new XElement(Ncx + "content", new XAttribute("src", NavigationHref(path, node, pendingPath))),
                BuildNcxNodes(node.Children, path, depth + 1, ref order, pendingPath)));
        }
        return result;
    }
    private string NavigationHref(string owner, EpubNavigationEntry node, string? pendingPath = null) {
        RequireText(node.Label, nameof(node.Label));
        EpubReference target = EpubReference.Resolve("package.opf", "/" + node.Target);
        if (target.Kind != EpubReferenceKind.Container || target.ContainerPath == null || (!_entries.ContainsKey(target.ContainerPath) && target.ContainerPath != pendingPath))
            throw new InvalidDataException("Navigation target must be a retained container resource: " + node.Target);
        return RelativeHref(owner, target.ContainerPath) + (target.Query == null ? string.Empty : "?" + target.Query) +
            (target.Fragment == null ? string.Empty : "#" + Uri.EscapeDataString(target.Fragment));
    }
    private static int NavigationDepth(IEnumerable<EpubNavigationEntry> nodes) =>
        nodes.Any() ? 1 + nodes.Max(node => NavigationDepth(node.Children)) : 0;

    private static void NormalizeNcxPlayOrder(XDocument navigation, string path) {
        var targets = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (XElement node in navigation.Descendants().Where(element => element.Name == Ncx + "navPoint" ||
            element.Name == Ncx + "pageTarget" || element.Name == Ncx + "navTarget")) {
            string source = (string?)node.Element(Ncx + "content")?.Attribute("src") ?? throw new InvalidDataException("NCX target has no source.");
            EpubReference reference = EpubReference.Resolve(path, source);
            string key = reference.Kind == EpubReferenceKind.Container ? reference.ContainerPath + "\0" + reference.Query + "\0" + reference.Fragment : source;
            if (!targets.TryGetValue(key, out int order)) { order = targets.Count + 1; targets.Add(key, order); }
            node.SetAttributeValue("playOrder", order);
        }
    }
}
