namespace OfficeIMO.Xps;

// Native structure owns names and reading order; consumers never rediscover it from paint.
internal sealed class XpsStoryFragmentsReader {
    internal sealed class Budget {
        private int _work, _characters;
        internal Budget(CancellationToken token) => Token = token;
        internal CancellationToken Token { get; }
        internal Dictionary<string, XElement> StoryParts { get; } = new(StringComparer.OrdinalIgnoreCase);
        internal void Charge(int count = 1) {
            Token.ThrowIfCancellationRequested();
            if (count > 1_000_000 - _work) throw new InvalidDataException("XPS content-structure work limit exceeded.");
            _work += count;
        }
        internal void Text(int count) {
            Charge();
            if (count > 16_000_000 - _characters) throw new InvalidDataException("XPS content-structure text limit exceeded.");
            _characters += count;
        }
    }
    private readonly XpsPage _page;
    private readonly int _pageIndex;
    private readonly Budget _budget;
    private readonly XNamespace _ns;
    private readonly List<string> _diagnostics = new();
    private readonly Dictionary<string, XElement> _names = new(StringComparer.Ordinal);
    private readonly Dictionary<XElement, int> _glyphs = new();
    private readonly Dictionary<string, XpsNamedContent?> _resolved = new(StringComparer.Ordinal);
    private bool _hasUnresolvedNames;

    internal XpsStoryFragmentsReader(XpsPage page, int pageIndex, XElement pageMarkup, Budget budget) {
        _page = page; _pageIndex = pageIndex; _budget = budget; _ns = page.Document.StructureNamespace;
        foreach (var element in PageElements(pageMarkup)) {
            _budget.Charge();
            if (element.Name.LocalName == "Glyphs") _glyphs.Add(element, _glyphs.Count);
            if ((string?)element.Attribute("Name") is string name) {
                if (_names.ContainsKey(name)) throw new InvalidDataException("Duplicate native page name: " + name);
                _names.Add(name, element);
            }
        }
    }
    internal static IEnumerable<XElement> PageElements(XElement root) {
        yield return root;
        if (root.Name.LocalName != "FixedPage" && root.Name.LocalName != "Canvas") yield break;
        foreach (var child in root.Elements()) {
            if (child.Name.Namespace != root.Name.Namespace || !new[] { "Canvas", "Path", "Glyphs" }.Contains(child.Name.LocalName)) continue;
            foreach (var element in PageElements(child)) yield return element;
        }
    }
    internal XpsPageStructure Read(XElement markup) {
        if (markup.Name != _ns + "StoryFragments") throw new InvalidDataException("Invalid StoryFragments root or dialect.");
        Attributes(markup, "");
        var fragments = new List<XpsStoryFragment>();
        var fragmentNames = new HashSet<(string? Story, string Fragment)>();
        foreach (var fragment in markup.Elements()) {
            _budget.Charge();
            if (fragment.Name != _ns + "StoryFragment") { Unknown(fragment); continue; }
            Attributes(fragment, "StoryName FragmentName FragmentType");
            string? name = (string?)fragment.Attribute("FragmentName");
            if (name != null && !fragmentNames.Add(((string?)fragment.Attribute("StoryName"), name))) throw new InvalidDataException("Duplicate name within a story's fragments.");
            string? type = (string?)fragment.Attribute("FragmentType");
            if (!Enum.TryParse(type, out XpsStoryFragmentType kind) || !Enum.IsDefined(typeof(XpsStoryFragmentType), kind) || kind.ToString() != type)
                throw new InvalidDataException("Invalid or missing story-fragment type.");
            var children = fragment.Elements().ToArray();
            bool before = children.Length > 0 && children[0].Name == _ns + "StoryBreak";
            bool after = children.Length > 0 && children[children.Length - 1].Name == _ns + "StoryBreak";
            var blocks = new List<XpsStructureNode>();
            for (int i = 0; i < children.Length; i++) {
                var child = children[i];
                if (child.Name == _ns + "StoryBreak") {
                    if (i != 0 && i != children.Length - 1) throw new InvalidDataException("StoryBreak must be at a fragment boundary.");
                    Attributes(child, "");
                    if (child.HasElements) throw new InvalidDataException("StoryBreak cannot contain elements.");
                } else blocks.Add(ReadNode(child, null));
            }
            if (blocks.Count == 0) throw new InvalidDataException("StoryFragment requires content structure.");
            fragments.Add(new XpsStoryFragment(_page.PartName, _pageIndex, (string?)fragment.Attribute("StoryName"), name,
                kind, before, after, blocks.AsReadOnly()));
        }
        if (fragments.Count == 0 && _diagnostics.Count == 0) throw new InvalidDataException("StoryFragments requires at least one fragment.");
        return new XpsPageStructure(true, fragments.AsReadOnly(), _diagnostics.AsReadOnly(), _hasUnresolvedNames);
    }
    private XpsStructureNode ReadNode(XElement element, XpsStructureKind? parent) {
        _budget.Charge();
        if (element.Name.Namespace != _ns) { Unknown(element); return Node(XpsStructureKind.Unknown); }
        XpsStructureKind kind = element.Name.LocalName switch {
            "SectionStructure" => XpsStructureKind.Section, "ParagraphStructure" => XpsStructureKind.Paragraph,
            "TableStructure" => XpsStructureKind.Table, "TableRowGroupStructure" => XpsStructureKind.TableRowGroup,
            "TableRowStructure" => XpsStructureKind.TableRow, "TableCellStructure" => XpsStructureKind.TableCell,
            "ListStructure" => XpsStructureKind.List, "ListItemStructure" => XpsStructureKind.ListItem,
            "FigureStructure" => XpsStructureKind.Figure, "NamedElement" => XpsStructureKind.NamedElement,
            _ => XpsStructureKind.Unknown
        };
        if (kind == XpsStructureKind.Unknown) { Unknown(element); return Node(kind); }
        if (!Allowed(parent, kind)) throw new InvalidDataException("Invalid native content-structure nesting: " + element.Name.LocalName);
        Attributes(element, kind == XpsStructureKind.NamedElement ? "NameReference" : kind == XpsStructureKind.TableCell ? "RowSpan ColumnSpan" : kind == XpsStructureKind.ListItem ? "Marker" : "");
        if (kind == XpsStructureKind.NamedElement) {
            if (element.HasElements) throw new InvalidDataException("NamedElement cannot contain elements.");
            string name = (string?)element.Attribute("NameReference") ?? throw new InvalidDataException("Missing NameReference.");
            return new XpsStructureNode(kind, Array.Empty<XpsStructureNode>(), Resolve(name), nameReference: name);
        }
        var children = element.Elements().Select(child => ReadNode(child, kind)).ToList();
        if (children.Count == 0 && new[] { XpsStructureKind.Section, XpsStructureKind.Table, XpsStructureKind.TableRowGroup, XpsStructureKind.TableRow, XpsStructureKind.List }.Contains(kind))
            throw new InvalidDataException("Native structural container requires children.");
        XpsNamedContent? marker = kind == XpsStructureKind.ListItem && element.Attribute("Marker") is XAttribute m ? Resolve(m.Value) : null;
        return new XpsStructureNode(kind, children.AsReadOnly(), marker: marker,
            rowSpan: kind == XpsStructureKind.TableCell ? Positive((string?)element.Attribute("RowSpan")) : 1,
            columnSpan: kind == XpsStructureKind.TableCell ? Positive((string?)element.Attribute("ColumnSpan")) : 1);
    }
    private static XpsStructureNode Node(XpsStructureKind kind) => new(kind, Array.Empty<XpsStructureNode>());
    private static bool Allowed(XpsStructureKind? parent, XpsStructureKind child) {
        bool block = child == XpsStructureKind.Paragraph || child == XpsStructureKind.Table || child == XpsStructureKind.List || child == XpsStructureKind.Figure;
        return parent switch {
            null => block || child == XpsStructureKind.Section,
            XpsStructureKind.Section or XpsStructureKind.TableCell or XpsStructureKind.ListItem => block,
            XpsStructureKind.Paragraph or XpsStructureKind.Figure => child == XpsStructureKind.NamedElement,
            XpsStructureKind.Table => child == XpsStructureKind.TableRowGroup,
            XpsStructureKind.TableRowGroup => child == XpsStructureKind.TableRow,
            XpsStructureKind.TableRow => child == XpsStructureKind.TableCell,
            XpsStructureKind.List => child == XpsStructureKind.ListItem,
            _ => false
        };
    }
    private XpsNamedContent? Resolve(string name) {
        _budget.Charge();
        try { XmlConvert.VerifyNCName(name); } catch (XmlException error) { throw new InvalidDataException("Invalid content name reference.", error); }
        if (_resolved.TryGetValue(name, out var cached)) {
            if (cached != null) _budget.Text(cached.Text.Length);
            return cached;
        }
        if (!_names.TryGetValue(name, out var target)) { _hasUnresolvedNames = true; Diagnostic("Unresolved native name: " + name); _resolved.Add(name, null); return null; }
        var ordinals = new List<int>(); var text = new StringBuilder();
        foreach (var element in PageElements(target)) {
            _budget.Charge();
            if (!_glyphs.TryGetValue(element, out int ordinal)) continue;
            ordinals.Add(ordinal);
            if (element.Attribute("UnicodeString") is XAttribute unicode) {
                string value = XpsPage.Unescape(unicode.Value); _budget.Text(value.Length); text.Append(value);
            } else Diagnostic("Referenced glyphs have no UnicodeString: " + name);
        }
        var content = new XpsNamedContent(name, target.Name.LocalName, _page.PartName, _pageIndex, text.ToString(), ordinals.AsReadOnly());
        _resolved.Add(name, content); return content;
    }
    private static int Positive(string? value) {
        if (value == null) return 1;
        if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int number) || number < 1)
            throw new InvalidDataException("Invalid table-cell span.");
        return number;
    }
    private void Attributes(XElement element, string allowed) {
        var names = allowed.Split(' ');
        foreach (var attribute in element.Attributes()) {
            _budget.Charge();
            if (attribute.IsNamespaceDeclaration || attribute.Name == XNamespace.Xml + "lang") continue;
            if (attribute.Name.NamespaceName.Length != 0 || !names.Contains(attribute.Name.LocalName))
                Diagnostic("Unsupported structure attribute: " + element.Name.LocalName + "." + attribute.Name);
        }
        if (element.Nodes().OfType<XText>().Any(t => !string.IsNullOrWhiteSpace(t.Value)))
            throw new InvalidDataException("Native content structure cannot contain literal text.");
    }
    private void Unknown(XElement element) => Diagnostic("Unsupported structure element: " + element.Name);
    private void Diagnostic(string value) { if (_diagnostics.Count < 100 && !_diagnostics.Contains(value)) _diagnostics.Add(value); }
}
