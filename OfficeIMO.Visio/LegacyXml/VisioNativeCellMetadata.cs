using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

/// <summary>
/// Bounded provenance for native cell state with no modern Cell attribute. Each model owns
/// independent, relative cell snapshots; document identity is assigned only by the writer.
/// </summary>
internal sealed class VisioNativeCellMetadata {
    internal static readonly XNamespace Namespace = "urn:officeimo:visio:legacy-cell-metadata:v1";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";
    private static readonly string[] ValueAttributes = { "N", "V", "F", "U", "E" };
    private readonly List<XElement> _entries;

    private VisioNativeCellMetadata(IEnumerable<XElement> entries) => _entries = entries.Select(e => new XElement(e)).ToList();
    internal VisioNativeCellMetadata Clone() => new(_entries);
    internal static VisioNativeCellMetadata Empty() => new(Array.Empty<XElement>());

    // Replacing artwork changes ownership even when a new row has the same cache.
    // Keep unrelated producer state on the instance and adopt the new definition's artwork state.
    internal void ReplaceArtworkState(VisioNativeCellMetadata? source, IReadOnlyDictionary<string, string>? references = null) {
        bool Artwork(XElement entry) {
            string address = (string?)entry.Attribute("Address") ?? string.Empty;
            return address.StartsWith("Section[N=Geometry]", StringComparison.Ordinal) ||
                address == "Cell[N=ImgOffsetX]" || address == "Cell[N=ImgOffsetY]" ||
                address == "Cell[N=ImgWidth]" || address == "Cell[N=ImgHeight]";
        }
        _entries.RemoveAll(entry => Artwork(entry));
        if (source == null) return;
        foreach (XElement entry in source._entries.Where(Artwork)) {
            var copy = new XElement(entry);
            if (references != null) foreach (XAttribute formula in copy.Elements(Namespace + "Snapshot").Attributes("F"))
                formula.Value = VisioShapeFormulaReferences.Rewrite(formula.Value, references)!;
            _entries.Add(copy);
        }
    }

    // Input is already subject to the canonical package XML byte/character and DTD limits.
    // Addresses are opaque keys, never XPath expressions or separately parsed XML payloads.
    internal static Dictionary<string, XElement?> Read(IEnumerable<XElement> elements) {
        var entries = new Dictionary<string, XElement?>(StringComparer.Ordinal);
        foreach (XElement entry in elements.Where(e => e.Name == Namespace + "NativeCellValues").Elements(Namespace + "Cell")) {
            string? address = (string?)entry.Attribute("Address");
            if (address == null || entry.Element(Namespace + "Snapshot") == null ||
                entry.Attribute("Condition") == null && entry.Attribute("Error") == null && entry.Element(Namespace + "LegacyFont") == null) continue;
            if (entries.ContainsKey(address)) entries[address] = null;
            else entries.Add(address, entry);
        }
        return entries;
    }

    internal static void BindTree(VisioShape model, XElement source, string scope, IReadOnlyDictionary<string, XElement?> entries) {
        model.NativeCellMetadata = Bind(source, scope, entries);
        XElement[] children = source.Element(Modern + "Shapes")?.Elements(Modern + "Shape").ToArray() ?? Array.Empty<XElement>();
        for (int i = 0; i < model.Children.Count && i < children.Length; i++) BindTree(model.Children[i], children[i], scope, entries);
    }

    // Additional master roots remain owned raw XML. Keep their own-cell provenance on the
    // master, keyed by durable native shape IDs, including descendants of each raw root.
    internal static void BindRawShapeTree(IDictionary<string, VisioNativeCellMetadata?> target,
        XElement source, string scope, IReadOnlyDictionary<string, XElement?> entries) {
        foreach (XElement shape in source.DescendantsAndSelf(Modern + "Shape")) {
            if ((string?)shape.Attribute("ID") is not string id) continue;
            id = VisioShapeFormulaReferences.NormalizeSheetId(id);
            VisioNativeCellMetadata? metadata = Bind(shape, scope, entries);
            if (target.ContainsKey(id)) target[id] = null;
            else target.Add(id, metadata);
        }
    }

    internal static VisioNativeCellMetadata? Bind(XElement source, string scope, IReadOnlyDictionary<string, XElement?> entries) {
        var owned = new List<XElement>();
        foreach (var pair in OwnCells(source)) {
            if (pair.Value == null) continue;
            if (!entries.TryGetValue(scope + "/" + Address(pair.Value), out XElement? entry) || entry == null) continue;
            // Claim the address even when stale: it must no longer match another object after copying.
            entry.Remove();
            if (!Matches(entry, pair.Value)) continue;
            var copy = new XElement(entry); copy.SetAttributeValue("Address", pair.Key); owned.Add(copy);
        }
        return owned.Count == 0 ? null : new VisioNativeCellMetadata(owned);
    }

    /// <summary>Rejects pre-copy edits before a mechanical formula rewrite can hide a difference.</summary>
    internal VisioNativeCellMetadata? Validated(XElement emitted, ISet<string>? assignedValues = null) {
        var valid = CurrentEntries(emitted, assignedValues).Select(pair => pair.Entry).ToArray();
        return valid.Length == 0 ? null : new VisioNativeCellMetadata(valid);
    }

    // Installing detached typed sections consumes their explicit value assignments.
    // Producer errors are independent of null conditions and remain snapshot guarded.
    internal void ForgetNullConditions(IEnumerable<string> assignedValues) {
        var addresses = new HashSet<string>(assignedValues, StringComparer.Ordinal);
        foreach (XElement entry in _entries)
            if (addresses.Contains((string)entry.Attribute("Address")!)) entry.Attribute("Condition")?.Remove();
        _entries.RemoveAll(entry => entry.Attribute("Condition") == null && entry.Attribute("Error") == null && entry.Element(Namespace + "LegacyFont") == null);
    }

    // A geometry row conversion can change a cell's meaning even when its numeric
    // cache is identical. Producer state belongs to the old native cell in that case.
    internal void ForgetCellState(XElement cell) {
        string address = string.Join("/", cell.AncestorsAndSelf().Reverse().Select(Step));
        _entries.RemoveAll(entry => (string?)entry.Attribute("Address") == address);
    }

    internal void RemapFormulas(IReadOnlyDictionary<string, string> ids) {
        foreach (XAttribute formula in _entries.Elements(Namespace + "Snapshot").Attributes("F"))
            formula.Value = VisioShapeFormulaReferences.Rewrite(formula.Value, ids)!;
    }

    internal void Collect(XElement emitted, string scope, XElement target, ISet<string>? assignedValues = null) {
        foreach (var pair in CurrentEntries(emitted, assignedValues)) {
            pair.Entry.SetAttributeValue("Address", scope + "/" + Address(pair.Cell));
            target.Add(pair.Entry);
        }
    }

    private IEnumerable<(XElement Entry, XElement Cell)> CurrentEntries(XElement emitted, ISet<string>? assignedValues) {
        var cells = OwnCells(emitted);
        foreach (XElement entry in _entries) {
            string address = (string)entry.Attribute("Address")!;
            if (!cells.TryGetValue(address, out XElement? cell) || cell == null || !Matches(entry, cell)) continue;
            // Typed value setters replace the native cell, including retained producer
            // errors whose canonical cache happens to equal the explicit assignment.
            if (assignedValues?.Contains(address) == true && VisioNativeCellAssignments.ReplacesProducerState(address)) continue;
            var copy = new XElement(entry);
            if (assignedValues?.Contains(address) == true) copy.Attribute("Condition")?.Remove();
            if (copy.Attribute("Condition") != null || copy.Attribute("Error") != null || copy.Element(Namespace + "LegacyFont") != null) yield return (copy, cell);
        }
    }

    private static Dictionary<string, XElement?> OwnCells(XElement shape) {
        var cells = new Dictionary<string, XElement?>(StringComparer.Ordinal);
        foreach (XElement cell in shape.Descendants(Modern + "Cell").Where(c =>
            shape.Name == Modern + "Shape" ? ReferenceEquals(c.Ancestors(Modern + "Shape").FirstOrDefault(), shape) : !c.Ancestors(Modern + "Shape").Any())) {
            string key = string.Join("/", cell.AncestorsAndSelf().TakeWhile(e => !ReferenceEquals(e, shape)).Reverse().Select(Step));
            if (cells.ContainsKey(key)) cells[key] = null;
            else cells.Add(key, cell);
        }
        return cells;
    }

    internal static bool Matches(XElement entry, XElement cell) => XNode.DeepEquals(Snapshot(cell), entry.Element(Namespace + "Snapshot"));
    internal static XElement Snapshot(XElement cell) => new(Namespace + "Snapshot", ValueAttributes.Where(n => cell.Attribute(n) != null).Select(n => new XAttribute(cell.Attribute(n)!)));
    internal static string Address(XElement element) => string.Join("/", element.AncestorsAndSelf().Reverse().Skip(1).Select(Step));
    internal static string MasterScope(string name) => "Masters/Master[NameU=" + Uri.EscapeDataString(name) + "]";
    internal static string PageScope(int id) => "Pages/Page[ID=" + id.ToString(CultureInfo.InvariantCulture) + "]";
    internal static string Step(XElement element) {
        string step = element.Name.LocalName;
        if (element.Name == Modern + "Master") return "Master[NameU=" + Uri.EscapeDataString((string?)element.Attribute("NameU") ?? string.Empty) + "]";
        foreach (string name in new[] { "ID", "N", "IX" }) {
            if (element.Attribute(name) is not XAttribute attribute) continue;
            string value = attribute.Value;
            if ((name == "ID" || name == "IX") && uint.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint index)) value = index.ToString(CultureInfo.InvariantCulture);
            step += "[" + name + "=" + Uri.EscapeDataString(value) + "]";
        }
        return step;
    }
}
