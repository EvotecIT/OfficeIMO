namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static readonly XNamespace Ncx = "http://www.daisy.org/z3986/2005/ncx/";

    /// <summary>Replaces TOC entries and optionally page-list/landmark entries while retaining unrelated navigation XML.</summary>
    public void SetNavigation(IEnumerable<EpubNavigationEntry> tableOfContents,
        IEnumerable<EpubNavigationEntry>? pageList = null, IEnumerable<EpubNavigationEntry>? landmarks = null) {
        if (tableOfContents == null) throw new ArgumentNullException(nameof(tableOfContents));
        EpubNavigationEntry[] toc = tableOfContents.ToArray();
        if (toc.Length == 0) throw new InvalidDataException("A nonempty table of contents is required.");
        EpubNavigationEntry[]? pages = pageList?.ToArray();
        EpubNavigationEntry[]? guide = landmarks?.ToArray();
        if (PackageVersion == "3.0" && guide != null) ValidateLandmarkTypes(guide, 0);
        if (PackageVersion == "2.0" && (pages?.Any(node => node.Children.Count != 0) == true || guide?.Any(node => node.Children.Count != 0) == true))
            throw new NotSupportedException("EPUB 2 page-list and guide entries must be flat; nested entries cannot be retained in these sections.");
        string path = NavigationPath();
        XDocument navigation = ParseXml(_entries[path], 64L * 1024 * 1024);
        ValidateNavigationRoot(navigation);
        XElement? newGuide = null;
        if (PackageVersion == "3.0") {
            XElement body = navigation.Root?.Element(Html + "body") ?? throw new InvalidDataException("Navigation has no XHTML body.");
            string linkOwner = HtmlNavigationLinkOwner(navigation, path);
            SetHtmlNavigation(body, "toc", "Contents", toc, linkOwner);
            if (pages != null) SetHtmlNavigation(body, "page-list", "Pages", pages, linkOwner);
            if (guide != null) SetHtmlNavigation(body, "landmarks", "Landmarks", guide, linkOwner);
        } else {
            XElement root = navigation.Root ?? throw new InvalidDataException("NCX has no root.");
            XElement map = root.Element(Ncx + "navMap") ?? throw new InvalidDataException("NCX has no navMap.");
            HashSet<string> ids = NavigationIds(navigation);
            int order = 0;
            ReplaceNavigationChildren(map, Ncx + "navPoint", BuildNcxNodes(toc, path, 0, ref order, ids));
            XElement? depth = root.Element(Ncx + "head")?.Elements(Ncx + "meta").FirstOrDefault(item => (string?)item.Attribute("name") == "dtb:depth");
            depth?.SetAttributeValue("content", NavigationDepth(toc).ToString(System.Globalization.CultureInfo.InvariantCulture));
            if (pages != null) {
                XElement? old = root.Element(Ncx + "pageList");
                if (pages.Length == 0) {
                    // An NCX pageList requires page targets, so an empty retained
                    // header/extension shell cannot be represented without loss.
                    if (old != null && (old.HasAttributes || old.Elements().Any(child => child.Name != Ncx + "pageTarget")))
                        throw new NotSupportedException("Clearing retained NCX page-list headers or extensions requires explicit content XML editing.");
                    old?.Remove();
                }
                else {
                    XElement list = old ?? new XElement(Ncx + "pageList", new XElement(Ncx + "navLabel", new XElement(Ncx + "text", "Pages")));
                    ReplaceNavigationChildren(list, Ncx + "pageTarget", pages.Select((page, index) => new XElement(Ncx + "pageTarget", new XAttribute("id", AllocateNavigationId(ids, "page-", index + 1)),
                        new XAttribute("playOrder", order + index + 1),
                        new XAttribute("type", page.SemanticType ?? "normal"), new XAttribute("value", index + 1),
                        new XElement(Ncx + "navLabel", new XElement(Ncx + "text", page.Label)),
                        new XElement(Ncx + "content", new XAttribute("src", NavigationHref(path, page))))));
                    if (old == null) map.AddAfterSelf(list);
                }
                foreach (string name in new[] { "dtb:totalPageCount", "dtb:maxPageNumber" }) {
                    XElement? count = root.Element(Ncx + "head")?.Elements(Ncx + "meta").FirstOrDefault(meta => (string?)meta.Attribute("name") == name);
                    count?.SetAttributeValue("content", pages.Length.ToString(System.Globalization.CultureInfo.InvariantCulture));
                }
            }
            NormalizeNcxPlayOrder(navigation, path);
            if (guide != null) {
                if (guide.Length != 0 || Root.Element(Opf + "guide") != null) {
                    newGuide = Root.Element(Opf + "guide") is XElement retained ? new XElement(retained) : new XElement(Opf + "guide");
                    ReplaceNavigationChildren(newGuide, Opf + "reference", guide.Select(item => new XElement(Opf + "reference",
                        new XAttribute("type", item.SemanticType ?? "text"), new XAttribute("title", item.Label),
                        new XAttribute("href", NavigationHref(PackagePath, item)))));
                    if (!newGuide.HasAttributes && !newGuide.Nodes().Any()) newGuide = null;
                }
            }
        }
        byte[] payload = SerializeXml(navigation);
        if (_encryption.Any(encryption => encryption.Path == path)) throw new NotSupportedException("Encrypted navigation cannot be edited.");
        if (payload.LongLength > _maximumEntryBytes) throw new InvalidDataException("Navigation exceeds the retained entry-byte limit.");
        long delta = payload.LongLength - _entries[path].LongLength;
        if (PackageVersion == "2.0" && guide != null) {
            EditPackageElement(Root, proposed => {
                XElement? old = proposed.Element(Opf + "guide");
                if (old != null && newGuide != null) old.ReplaceWith(new XElement(newGuide));
                else if (old != null) old.Remove();
                else if (newGuide != null) proposed.Add(new XElement(newGuide));
            }, delta);
        } else EnsurePackageBudget(_package, delta);
        _entries[path] = payload;
        _retainedBytes += delta;
        MarkChanged();
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
            EditPackageElement(RequireSection("spine"), proposed => proposed.SetAttributeValue("toc", "navigation"));
        }
    }
    private string NavigationPath() {
        EpubManifestItem? item = PackageVersion == "3.0" ? Manifest.SingleOrDefault(resource => HasToken(resource.Properties, "nav")) :
            Manifest.SingleOrDefault(resource => resource.Id == (string?)RequireSection("spine").Attribute("toc"));
        if (item == null) throw new InvalidDataException("Package has no declared navigation resource.");
        string expected = PackageVersion == "3.0" ? "application/xhtml+xml" : "application/x-dtbncx+xml";
        if (!HasMediaType(item.MediaType, expected)) throw new InvalidDataException("Navigation resource must declare " + expected + ".");
        string path = RequireLocalPath(item);
        if (!_entries.ContainsKey(path)) throw new InvalidDataException("Navigation resource is missing.");
        return path;
    }
    private byte[] PrepareAppendedNavigation(EpubNavigationEntry entry, string pendingPath) {
        string path = NavigationPath();
        XDocument navigation = ParseXml(_entries[path], 64L * 1024 * 1024);
        ValidateNavigationRoot(navigation);
        if (PackageVersion == "3.0") {
            XElement nav = navigation.Descendants(Html + "nav").Single(element => HasToken((string?)element.Attribute(Ops + "type"), "toc"));
            XElement list = nav.Element(Html + "ol") ?? throw new InvalidDataException("TOC list is missing.");
            list.Add(BuildHtmlNodes(new[] { entry }, HtmlNavigationLinkOwner(navigation, path), 0, pendingPath));
        } else {
            XElement map = navigation.Root?.Element(Ncx + "navMap") ?? throw new InvalidDataException("NCX navMap is missing.");
            int order = map.Descendants(Ncx + "navPoint").Count();
            map.Add(BuildNcxNodes(new[] { entry }, path, 0, ref order, NavigationIds(navigation), pendingPath));
            NormalizeNcxPlayOrder(navigation, path);
        }
        return SerializeXml(navigation);
    }
    private void ValidateNavigationRoot(XDocument navigation) {
        if (navigation.Root?.Name != (PackageVersion == "3.0" ? Html + "html" : Ncx + "ncx"))
            throw new InvalidDataException("Navigation root does not match the package generation.");
    }
    private static string HtmlNavigationLinkOwner(XDocument navigation, string path) {
        string? baseHref = navigation.Root?.Element(Html + "head")?.Elements(Html + "base")
            .Select(element => (string?)element.Attribute("href")).FirstOrDefault(value => value != null);
        if (string.IsNullOrWhiteSpace(baseHref)) return path;
        EpubReference directory = EpubReference.Resolve(path, baseHref, ".");
        if (directory.Kind == EpubReferenceKind.External)
            throw new NotSupportedException("Container navigation cannot be generated under an external HTML base URL.");
        if (directory.Kind != EpubReferenceKind.Container || directory.ContainerPath == null)
            throw new InvalidDataException("Navigation has an invalid HTML base URL.");
        // Use the shared resolver to determine the effective directory, including
        // directory/file bases and encoded paths, before computing relative links.
        return (directory.ContainerPath.Length == 0 ? string.Empty : directory.ContainerPath + "/") + "__officeimo_navigation_base__";
    }
    private void SetHtmlNavigation(XElement body, string type, string heading, EpubNavigationEntry[] nodes, string path) {
        XElement? nav = body.Descendants(Html + "nav").FirstOrDefault(element => HasToken((string?)element.Attribute(Ops + "type"), type));
        if (nodes.Length == 0 && type != "toc") {
            if (nav != null) {
                // Clearing an optional section retains extension attributes/children;
                // removing its semantic type makes the remaining shell ordinary XHTML.
                nav.SetAttributeValue(Ops + "type", NormalizeProperties(string.Join(" ", Tokens((string?)nav.Attribute(Ops + "type")).Where(token => token != type))));
                XElement? list = nav.Element(Html + "ol");
                if (list != null) ReplaceNavigationChildren(list, Html + "li", Array.Empty<XElement>());
            }
            return;
        }
        if (nav == null) {
            nav = new XElement(Html + "nav", new XAttribute(Ops + "type", type), new XElement(Html + "h1", heading));
            body.Add(nav);
        }
        XElement? old = nav.Element(Html + "ol");
        if (old != null) ReplaceNavigationChildren(old, Html + "li", BuildHtmlNodes(nodes, path, 0));
        else nav.Add(new XElement(Html + "ol", BuildHtmlNodes(nodes, path, 0)));
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
    private IEnumerable<XElement> BuildNcxNodes(IEnumerable<EpubNavigationEntry> nodes, string path, int depth, ref int order, HashSet<string> ids, string? pendingPath = null) {
        if (depth > 64) throw new InvalidDataException("Navigation depth exceeds 64.");
        var result = new List<XElement>();
        foreach (EpubNavigationEntry node in nodes) {
            int current = ++order;
            result.Add(new XElement(Ncx + "navPoint", new XAttribute("id", AllocateNavigationId(ids, "nav-", current)), new XAttribute("playOrder", current),
                new XElement(Ncx + "navLabel", new XElement(Ncx + "text", node.Label)),
                new XElement(Ncx + "content", new XAttribute("src", NavigationHref(path, node, pendingPath))),
                BuildNcxNodes(node.Children, path, depth + 1, ref order, ids, pendingPath)));
        }
        return result;
    }
    private string NavigationHref(string owner, EpubNavigationEntry node, string? pendingPath = null) {
        RequireText(node.Label, nameof(node.Label));
        EpubReference target = EpubReference.Resolve("package.opf", "/" + node.Target);
        if (target.Kind != EpubReferenceKind.Container || target.ContainerPath == null || (!_entries.ContainsKey(target.ContainerPath) && target.ContainerPath != pendingPath))
            throw new InvalidDataException("Navigation target must be a retained container resource: " + node.Target);
        if (target.ContainerPath != pendingPath)
            RequireSpineTarget(target, Manifest, Spine.Select(item => item.ManifestId));
        return RelativeHref(owner, target.ContainerPath) + (target.Query == null ? string.Empty : "?" + target.Query) +
            (target.Fragment == null ? string.Empty : "#" + Uri.EscapeDataString(target.Fragment));
    }
    private static void ValidateLandmarkTypes(IEnumerable<EpubNavigationEntry> nodes, int depth) {
        if (depth > 64) throw new InvalidDataException("Navigation depth exceeds 64.");
        foreach (EpubNavigationEntry node in nodes) {
            if (string.IsNullOrWhiteSpace(node.SemanticType)) throw new InvalidDataException("Every landmark link requires a semantic type.");
            ValidateLandmarkTypes(node.Children, depth + 1);
        }
    }

    private static void RequireSpineTarget(EpubReference target, IEnumerable<EpubManifestItem> manifest, IEnumerable<string> spineIds, bool requireContainer = false) {
        if (target.Kind != EpubReferenceKind.Container) {
            if (requireContainer) throw new InvalidDataException("Navigation must target container content: " + target.Original);
            return;
        }
        var ids = new HashSet<string>(spineIds, StringComparer.Ordinal);
        if (!manifest.Any(item => item.Reference.ContainerPath == target.ContainerPath && ids.Contains(item.Id)))
            throw new InvalidDataException("Navigation content target must be declared in the manifest and spine: " + target.Original);
    }
    private static int NavigationDepth(IEnumerable<EpubNavigationEntry> nodes) =>
        nodes.Any() ? 1 + nodes.Max(node => NavigationDepth(node.Children)) : 0;

    private static HashSet<string> NavigationIds(XDocument document) => new HashSet<string>(document.Descendants().Attributes()
        .Where(attribute => attribute.Name == "id" || attribute.Name == XNamespace.Xml + "id").Select(attribute => attribute.Value), StringComparer.Ordinal);

    private static string AllocateNavigationId(HashSet<string> ids, string prefix, int start) {
        string candidate;
        do { candidate = prefix + start++.ToString(System.Globalization.CultureInfo.InvariantCulture); } while (!ids.Add(candidate));
        return candidate;
    }

    private static void ReplaceNavigationChildren(XElement parent, XName ownedName, IEnumerable<XElement> replacements) {
        XElement[] added = replacements.ToArray();
        XElement[] old = parent.Elements(ownedName).ToArray();
        if (old.Length == 0) parent.Add(added); else old[0].AddBeforeSelf(added);
        foreach (XElement child in old) child.Remove();
    }

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
