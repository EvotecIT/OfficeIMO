namespace OfficeIMO.OpenDocument;

// One counter walk per text story. Unsupported labels do not discard the item's editable body.
internal sealed class OdfDrawingListResolver {
    internal sealed class Entry {
        internal Entry(XElement level, string? label) { Level = level; Label = label; }
        internal XElement Level { get; }
        internal string? Label { get; }
    }
    private sealed class Counter {
        internal Counter(XElement? style, int level, long[] values, bool valid = true) { Style = style; Level = level; Values = values; Valid = valid; }
        internal XElement? Style { get; }
        internal int Level { get; }
        internal long[] Values { get; }
        internal bool Valid { get; }
    }
    private readonly OdgShape _shape;
    private readonly HashSet<string> _losses;
    private readonly Dictionary<XElement, Entry> _entries = new Dictionary<XElement, Entry>();
    private readonly Dictionary<string, XElement?> _styles = new Dictionary<string, XElement?>(StringComparer.Ordinal);
    private readonly Dictionary<string, Counter> _ids = new Dictionary<string, Counter>(StringComparer.Ordinal);
    private readonly HashSet<string> _ambiguousIds = new HashSet<string>(StringComparer.Ordinal);
    private readonly Dictionary<(XElement, int), Counter> _previous = new Dictionary<(XElement, int), Counter>();
    private readonly XElement?[] _activeLevels = new XElement?[129];
    private int _visited;

    internal OdfDrawingListResolver(OdgShape shape, HashSet<string> losses) {
        _shape = shape; _losses = losses;
        IndexStyles(shape.Document.GetXml(shape.PartPath).Root?.Element(OdfNamespaces.Office + "automatic-styles"), overwrite: true);
        if (shape.Document.Package.ContainsEntry("styles.xml"))
            IndexStyles(shape.Document.GetXml("styles.xml").Root?.Element(OdfNamespaces.Office + "styles"), overwrite: false);
        // Duplicate IDs invalidate continuation even when a duplicate occurs later in the story.
        foreach (XElement element in shape.TextRoot.Descendants(OdfNamespaces.Text + "list")) {
            if (++_visited > 100000) throw new NotSupportedException("List traversal exceeds the projection node limit.");
            string? id = (string?)element.Attribute(XNamespace.Xml + "id");
            if (id != null && !_ambiguousIds.Add(id)) _ids[id] = new Counter(null, 0, Array.Empty<long>());
        }
        _ambiguousIds.Clear();
        foreach (string id in _ids.Keys) _ambiguousIds.Add(id);
        _ids.Clear(); _visited = 0;
        Walk(shape.TextRoot, 0, null, new long[129], 0);
    }

    internal Entry? For(XElement paragraph) => _entries.TryGetValue(paragraph, out Entry? entry) ? entry : null;

    private void IndexStyles(XElement? container, bool overwrite) {
        if (container == null) return;
        foreach (IGrouping<string, XElement> group in container.Elements(OdfNamespaces.Text + "list-style")
            .Where(e => e.Attribute(OdfNamespaces.Style + "name") != null).GroupBy(e => (string)e.Attribute(OdfNamespaces.Style + "name")!, StringComparer.Ordinal)) {
            if (!overwrite && _styles.ContainsKey(group.Key)) continue;
            _styles[group.Key] = group.Count() == 1 ? group.First() : null;
        }
    }

    private void Walk(XElement container, int level, XElement? inheritedStyle, long[] ancestors, int depth) {
        if (depth > 128) throw new NotSupportedException("List nesting exceeds the projection limit.");
        foreach (XElement child in container.Elements()) {
            if (++_visited > 100000) throw new NotSupportedException("List traversal exceeds the projection node limit.");
            if (child.Name != OdfNamespaces.Text + "list") continue;
            ReadList(child, level + 1, inheritedStyle, ancestors, depth + 1);
        }
    }

    private void ReadList(XElement list, int level, XElement? inheritedStyle, long[] ancestors, int depth) {
        if (depth > 128 || level >= ancestors.Length) throw new NotSupportedException("List nesting exceeds the projection limit.");
        XElement? firstParagraph = list.Elements().SelectMany(e => e.Elements().Where(OdfTextTraversal.IsParagraph)).FirstOrDefault();
        var paragraph = firstParagraph == null ? null : new OdfTextParagraph(_shape.Document, firstParagraph, _shape.Element);
        string? styleName = (string?)list.Attribute(OdfNamespaces.Text + "style-name");
        XElement? style = styleName != null ? Find(styleName) : inheritedStyle ?? DefaultStyle(paragraph);
        bool validCounter = style != null;
        if (!validCounter) _losses.Add("list-style");
        if (IsTrue((string?)style?.Attribute(OdfNamespaces.Text + "consecutive-numbering"))) {
            _losses.Add("list-consecutive-numbering"); validCounter = false;
        }
        long[] values = (long[])ancestors.Clone();
        values[level] = 0;
        string? continuedId = (string?)list.Attribute(OdfNamespaces.Text + "continue-list");
        bool continueNumbering = IsTrue((string?)list.Attribute(OdfNamespaces.Text + "continue-numbering"));
        if (continuedId != null || continueNumbering) {
            Counter? previous = null;
            bool found = continuedId != null ? !_ambiguousIds.Contains(continuedId) && _ids.TryGetValue(continuedId, out previous) :
                style != null && _previous.TryGetValue((style, level), out previous);
            if (!found || previous == null || !previous.Valid || previous.Style != style || previous.Level != level) {
                _losses.Add("list-continuation"); validCounter = false;
            } else values[level] = previous.Values[level];
        }
        XElement? defaultLevel = Level(style, level);
        bool firstItem = true;
        foreach (XElement item in list.Elements()) {
            if (++_visited > 100000) throw new NotSupportedException("List traversal exceeds the projection node limit.");
            bool isItem = item.Name == OdfNamespaces.Text + "list-item", header = item.Name == OdfNamespaces.Text + "list-header";
            if (!isItem && !header) { _losses.Add("list-container"); continue; }
            XElement? levelStyle = defaultLevel;
            string? overrideName = (string?)item.Attribute(OdfNamespaces.Text + "style-override");
            if (overrideName != null) levelStyle = Level(Find(overrideName), level);
            bool validItem = validCounter && levelStyle != null;
            if (levelStyle == null) _losses.Add("list-level-style");
            string? label = null;
            bool counterUpdated = false;
            try {
                _activeLevels[level] = levelStyle;
                if (isItem) {
                    string? start = (string?)item.Attribute(OdfNamespaces.Text + "start-value");
                    if (start != null) values[level] = Number(start);
                    else if (firstItem && !continueNumbering && continuedId == null) values[level] = Number((string?)levelStyle?.Attribute(OdfNamespaces.Text + "start-value") ?? "1");
                    else values[level] = checked(values[level] + 1);
                    firstItem = false;
                    counterUpdated = true;
                    if (validItem) label = Label(levelStyle!, _activeLevels, values, level);
                }
            } catch (Exception exception) when (exception is NotSupportedException or OverflowException or FormatException) {
                _losses.Add("list-numbering"); validItem = false;
                if (!counterUpdated) validCounter = false;
            }
            bool first = true;
            foreach (XElement p in item.Elements().Where(OdfTextTraversal.IsParagraph)) {
                if (levelStyle != null) _entries[p] = new Entry(levelStyle, first && isItem && validItem ? label : null);
                first = false;
            }
            Walk(item, level, style, values, depth + 1);
        }
        string? id = (string?)list.Attribute(XNamespace.Xml + "id");
        var state = new Counter(style, level, values, validCounter);
        if (id != null && !_ambiguousIds.Contains(id)) _ids[id] = state;
        if (style != null) _previous[(style, level)] = state;
    }

    private XElement? DefaultStyle(OdfTextParagraph? paragraph) {
        if (paragraph == null) return null;
        foreach (OdfStyle style in paragraph.Styles) {
            string? name = (string?)style.Element.Attribute(OdfNamespaces.Style + "list-style-name");
            if (name != null) return Find(name);
            XElement? embedded = style.Element.Element(OdfNamespaces.Style + "graphic-properties")?.Element(OdfNamespaces.Text + "list-style");
            if (embedded != null) return embedded;
        }
        return null;
    }
    private XElement? Find(string name) => _styles.TryGetValue(name, out XElement? style) ? style : null;
    private static XElement? Level(XElement? style, int level) {
        XElement[] matches = style?.Elements().Where(e => int.TryParse((string?)e.Attribute(OdfNamespaces.Text + "level"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int defined) && defined == level).Take(2).ToArray() ?? Array.Empty<XElement>();
        return matches.Length == 1 ? matches[0] : null;
    }
    private static long Number(string text) {
        if (!long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long value) || value < 0) throw new NotSupportedException("Unsupported list counter value.");
        return value;
    }
    private static string? Label(XElement style, XElement?[] activeLevels, long[] values, int level) {
        string prefix = (string?)style.Attribute(OdfNamespaces.Style + "num-prefix") ?? "", suffix = (string?)style.Attribute(OdfNamespaces.Style + "num-suffix") ?? "";
        if (prefix.Length + (long)suffix.Length > 4096 || style.Attribute(OdfNamespaces.Style + "num-list-format-name") != null)
            throw new NotSupportedException("List label length or named number-string format is outside the profile.");
        if (style.Name == OdfNamespaces.Text + "list-level-style-bullet") return prefix + ((string?)style.Attribute(OdfNamespaces.Text + "bullet-char") ?? throw new NotSupportedException("Missing bullet character.")) + suffix;
        if (style.Name != OdfNamespaces.Text + "list-level-style-number") throw new NotSupportedException("Image list labels are outside the text-label profile.");
        string format = (string?)style.Attribute(OdfNamespaces.Style + "num-format") ?? "";
        if (format.Length == 0) return null;
        long display = Number((string?)style.Attribute(OdfNamespaces.Text + "display-levels") ?? "1");
        if (display < 1 || display > level) throw new NotSupportedException("Invalid list display levels.");
        return prefix + string.Join(".", Enumerable.Range(level - (int)display + 1, (int)display).Select(i => {
            XElement component = activeLevels[i] ?? throw new NotSupportedException("Missing ancestor list number format.");
            if (component.Name != OdfNamespaces.Text + "list-level-style-number") throw new NotSupportedException("A displayed ancestor list level must be numbered.");
            return OdfNumberingFormatter.Format(values[i], (string?)component.Attribute(OdfNamespaces.Style + "num-format") ?? "", IsTrue((string?)component.Attribute(OdfNamespaces.Style + "num-letter-sync")));
        })) + suffix;
    }
    private static bool IsTrue(string? value) => value is "true" or "1";
}
