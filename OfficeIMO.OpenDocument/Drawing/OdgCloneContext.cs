namespace OfficeIMO.OpenDocument;

/// <summary>Plans native same-document copies before attaching any XML or changing package entries.</summary>
internal sealed class OdgCloneContext {
    private readonly HashSet<string> _ids;
    private readonly HashSet<string> _names;
    private readonly HashSet<string> _noteNames;
    private readonly HashSet<string> _listGroups;
    private readonly Dictionary<string, int> _next = new Dictionary<string, int>(StringComparer.Ordinal);
    private static readonly HashSet<XName> IdReferences = new HashSet<XName> {
        OdfNamespaces.Draw + "start-shape", OdfNamespaces.Draw + "end-shape", OdfNamespaces.Draw + "caption-id",
        OdfNamespaces.Draw + "shape-id", OdfNamespaces.Text + "continue-list",
        OdfNamespaces.Presentation + "master-element", OdfNamespaces.Smil + "targetElement", OdfNamespaces.Smil + "endsync"
    };

    internal OdgCloneContext(params XElement[] roots) {
        _ids = new HashSet<string>(roots.SelectMany(root => root.DescendantsAndSelf()).SelectMany(Identifiers).Select(attribute => attribute.Value), StringComparer.Ordinal);
        _names = new HashSet<string>(roots.SelectMany(root => root.DescendantsAndSelf()).Attributes(OdfNamespaces.Draw + "name").Select(attribute => attribute.Value), StringComparer.Ordinal);
        _noteNames = new HashSet<string>(roots.SelectMany(root => root.DescendantsAndSelf()).Where(element => element.Name == OdfNamespaces.Text + "note")
            .Attributes(OdfNamespaces.Text + "id").Select(attribute => attribute.Value), StringComparer.Ordinal);
        _listGroups = new HashSet<string>(roots.SelectMany(root => root.DescendantsAndSelf()).Attributes(OdfNamespaces.Text + "list-id").Select(attribute => attribute.Value), StringComparer.Ordinal);
    }

    internal XElement[] CloneForest(IEnumerable<XElement> sources, string pageName, string destinationPageName) =>
        Clone(new XElement("copy-plan", sources.Select(element => new XElement(element))), pageName, destinationPageName, true).Elements().ToArray();

    internal XElement Clone(XElement source, string? pageName = null, string? destinationPageName = null, bool closedReferences = false) {
        XElement clone = new XElement(source);
        if (destinationPageName != null) _names.Add(destinationPageName);
        var identifiers = new Dictionary<string, string>(StringComparer.Ordinal);
        var owners = new Dictionary<string, XElement>(StringComparer.Ordinal);
        var shapeNames = new Dictionary<string, string>(StringComparer.Ordinal);
        var ambiguousNames = new HashSet<string>(StringComparer.Ordinal);
        var notes = new Dictionary<string, string>(StringComparer.Ordinal);
        var lists = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (XElement element in clone.DescendantsAndSelf()) {
            RequireSupported(element);
            if (closedReferences && (element.Name.Namespace == OdfNamespaces.Script || element.Name == OdfNamespaces.Text + "script" || element.Name == OdfNamespaces.Office + "event-listeners"))
                throw new NotSupportedException("Drawing import does not copy scripts or event listeners.");
            if (element.Name == OdfNamespaces.Text + "note" && element.Attribute(OdfNamespaces.Text + "id") is XAttribute note) {
                if (notes.ContainsKey(note.Value)) throw new InvalidDataException("Cloned note name is ambiguous: " + note.Value);
                string replacement = Reserve(_noteNames, "odgCloneNote"); notes.Add(note.Value, replacement); note.Value = replacement;
            }
            if (element.Name == OdfNamespaces.Text + "numbered-paragraph" && element.Attribute(OdfNamespaces.Text + "list-id") is XAttribute list) {
                if (!lists.TryGetValue(list.Value, out string? replacement)) {
                    replacement = Reserve(_listGroups, "odgCloneList"); lists.Add(list.Value, replacement);
                }
                list.Value = replacement;
            }
            foreach (XAttribute attribute in Identifiers(element)) {
                string original = attribute.Value;
                if (owners.TryGetValue(original, out XElement? owner) && !ReferenceEquals(owner, element))
                    throw new InvalidDataException("Cloned content has an ambiguous identifier: " + original);
                owners[original] = element;
                if (!identifiers.TryGetValue(original, out string? replacement)) {
                    replacement = Reserve(_ids, "odgClone"); identifiers.Add(original, replacement);
                }
                attribute.Value = replacement;
            }
            XAttribute? name = element.Attribute(OdfNamespaces.Draw + "name");
            if (name == null || element.Name == OdfNamespaces.Draw + "layer") continue;
            if (pageName != null && element.Name == OdfNamespaces.Draw + "page" && name.Value == pageName) { name.Value = destinationPageName!; continue; }
            string oldName = name.Value;
            string newName = Reserve(_names, oldName + "Copy");
            if (shapeNames.ContainsKey(oldName)) ambiguousNames.Add(oldName);
            else shapeNames.Add(oldName, newName);
            name.Value = newName;
        }
        if (pageName != null && clone.Name == OdfNamespaces.Draw + "page") clone.SetAttributeValue(OdfNamespaces.Draw + "name", destinationPageName);
        foreach (XAttribute attribute in clone.DescendantsAndSelf().Attributes()) {
            if (IdReferences.Contains(attribute.Name)) {
                if (identifiers.TryGetValue(attribute.Value, out string? replacement)) attribute.Value = replacement;
                else if (closedReferences || attribute.Name == OdfNamespaces.Draw + "start-shape" || attribute.Name == OdfNamespaces.Draw + "end-shape")
                    throw new InvalidDataException("Cloned connector attachment is outside the copied content: " + attribute.Value);
            } else if (attribute.Name == OdfNamespaces.Text + "ref-name" && attribute.Parent?.Name == OdfNamespaces.Text + "note-ref") {
                if (notes.TryGetValue(attribute.Value, out string? replacement)) attribute.Value = replacement;
                else if (closedReferences) throw new InvalidDataException("Imported note reference is outside the copied content: " + attribute.Value);
            } else if (attribute.Name == OdfNamespaces.Draw + "nav-order") {
                attribute.Value = string.Join(" ", attribute.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
                    .Select(value => identifiers.TryGetValue(value, out string? replacement) ? replacement
                        : throw new InvalidDataException("Cloned navigation target is outside the copied content: " + value)));
            } else if (attribute.Name == OdfNamespaces.Draw + "chain-next-name") {
                if (ambiguousNames.Contains(attribute.Value) || !shapeNames.TryGetValue(attribute.Value, out string? replacement))
                    throw new InvalidDataException("Cloned text-box chain target is missing or ambiguous: " + attribute.Value);
                attribute.Value = replacement;
            } else if (attribute.Name == OdfNamespaces.XLink + "href" && attribute.Value.StartsWith("#", StringComparison.Ordinal)) {
                string target = Uri.UnescapeDataString(attribute.Value.Substring(1));
                if (identifiers.TryGetValue(target, out string? replacement)) attribute.Value = "#" + Uri.EscapeDataString(replacement);
                else if (target == pageName) attribute.Value = "#" + Uri.EscapeDataString(destinationPageName!);
                else if (shapeNames.TryGetValue(target, out replacement)) {
                    if (ambiguousNames.Contains(target)) throw new InvalidDataException("Cloned hyperlink target is ambiguous: " + target);
                    attribute.Value = "#" + Uri.EscapeDataString(replacement);
                }
                else if (closedReferences) throw new InvalidDataException("Imported fragment link is outside the copied content: " + target);
            }
        }
        return clone;
    }

    private static IEnumerable<XAttribute> Identifiers(XElement element) => element.Attributes().Where(attribute =>
        attribute.Name == XNamespace.Xml + "id" || attribute.Name == OdfNamespaces.Text + "id" && IsTextIdAlias(element) ||
        attribute.Name == OdfNamespaces.Draw + "id" && element.Name != OdfNamespaces.Draw + "glue-point");
    private static bool IsTextIdAlias(XElement element) => element.Name == OdfNamespaces.Text + "p" ||
        element.Name == OdfNamespaces.Text + "h" || element.Name == OdfNamespaces.Draw + "text-box";
    private string Reserve(HashSet<string> values, string prefix) {
        int index = _next.TryGetValue(prefix, out int next) ? next : 1; string name;
        do { name = prefix + index++.ToString(CultureInfo.InvariantCulture); } while (!values.Add(name));
        _next[prefix] = index;
        return name;
    }
    private static void RequireSupported(XElement element) {
        if (element.Name.Namespace == OdfNamespaces.Anim || element.Name == OdfNamespaces.Presentation + "animations" ||
            element.Name == OdfNamespaces.Office + "forms" || element.Name == OdfNamespaces.Text + "tracked-changes" ||
            element.Name == OdfNamespaces.Draw + "object" || element.Name == OdfNamespaces.Draw + "object-ole" ||
            element.Name == OdfNamespaces.Text + "sequence" ||
            element.Name == OdfNamespaces.Draw + "control" || element.Attributes().Any(attribute =>
                attribute.Name == OdfNamespaces.Text + "id" && !IsTextIdAlias(element) && element.Name != OdfNamespaces.Text + "note" ||
                attribute.Name == OdfNamespaces.Text + "name" || attribute.Name == OdfNamespaces.Office + "name" ||
                attribute.Name == OdfNamespaces.Table + "name" || attribute.Name == OdfNamespaces.Text + "change-id"))
            throw new NotSupportedException("Drawing cloning does not support embedded editable objects, forms, animations, tracked changes or named text/table definitions: " + element.Name.LocalName);
    }
}
