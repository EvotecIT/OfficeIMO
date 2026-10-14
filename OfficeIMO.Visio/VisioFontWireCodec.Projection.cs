using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

internal sealed partial class VisioFontWireCodec {
    /// <summary>Declares literal family names introduced through editable native ShapeSheet rows.</summary>
    internal static void DeclareLiteralNames(XElement document, IEnumerable<XElement> roots) {
        XElement? faces = document.Element(Modern + "FaceNames");
        var names = new HashSet<string>(faces?.Elements(Modern + "FaceName").Attributes("Name").Select(a => a.Value) ?? Enumerable.Empty<string>(), StringComparer.Ordinal);
        var ids = new HashSet<int>(faces?.Elements(Modern + "FaceName").Attributes("ID").Select(a =>
            int.TryParse(a.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id) ? id : -1) ?? Enumerable.Empty<int>());
        int next = 1;
        foreach (XElement cell in roots.SelectMany(FontCells)) {
            string? value = (string?)cell.Attribute("V");
            if (string.IsNullOrWhiteSpace(value) || int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _) ||
                string.Equals(value, "themed", StringComparison.OrdinalIgnoreCase) || !names.Add(value!)) continue;
            if (faces == null) {
                faces = new XElement(Modern + "FaceNames");
                XElement? styles = document.Element(Modern + "StyleSheets");
                if (styles == null) document.Add(faces); else styles.AddBeforeSelf(faces);
            }
            while (ids.Contains(next)) next++;
            faces.Add(new XElement(Modern + "FaceName", new XAttribute("ID", next), new XAttribute("Name", value!)));
            ids.Add(next++);
        }
    }

    /// <summary>
    /// Projects a finalized private output tree. Legacy-only table data lives in the existing
    /// document preservation namespace, never as undeclared native FaceName attributes.
    /// </summary>
    internal static VisioFontWireCodec WriteDocument(XElement root) {
        XElement? faces = root.Element(Modern + "FaceNames"), fonts = root.Element(Modern + "Fonts");
        var codec = new VisioFontWireCodec(faces);
        var native = new XElement(Modern + "FaceNames");
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (XElement face in faces?.Elements(Modern + "FaceName") ?? Enumerable.Empty<XElement>()) {
            string? name = (string?)face.Attribute("Name") ?? (string?)face.Attribute("NameU");
            if (string.IsNullOrWhiteSpace(name) || !names.Add(name!)) continue;
            native.Add(new XElement(Modern + "FaceName", new XAttribute("NameU", name!),
                FaceAttributes.Where(attribute => face.Attribute(attribute) != null).Select(attribute => new XAttribute(face.Attribute(attribute)!))));
        }
        var original = new XElement(Metadata + "Original");
        if (faces != null && (faces.HasElements || faces.HasAttributes)) original.Add(PackTable(faces));
        if (fonts != null) original.Add(PackTable(fonts));
        if (faces != null) { if (native.HasElements) faces.ReplaceWith(native); else faces.Remove(); }
        else if (native.HasElements) {
            XElement? following = root.Element(Modern + "StyleSheets");
            if (following == null) root.Add(native); else following.AddBeforeSelf(native);
        }
        fonts?.Remove();
        foreach (XElement old in root.Elements(Metadata + "FontTable").ToArray()) old.Remove();
        if (original.HasElements) root.Add(new XElement(Metadata + "FontTable", original,
            new XElement(Metadata + "Snapshot", native.HasElements ? PackTable(native) : null)));
        return codec;
    }

    /// <summary>Records the original numeric value/formula before emitting cached font names.</summary>
    internal void EncodeCells(XElement root, XElement documentRoot, IReadOnlyDictionary<string, XElement?> entries, string? scope = null) {
        XElement? added = null;
        foreach (XElement cell in FontCells(root)) {
            var original = VisioNativeCellMetadata.Snapshot(cell);
            string? value = (string?)cell.Attribute("V");
            if (!IsFallbackSentinel(cell, value) && value != null && _namesById.TryGetValue(Identity(value), out string? name)) cell.SetAttributeValue("V", name);
            if (!IsFallbackSentinel(cell, value) && cell.Attribute("F") is XAttribute formula)
                formula.Value = RewriteConstantFormula(formula.Value, id => _namesById.TryGetValue(id, out string? family)
                    ? "FONT(\"" + family.Replace("\"", "\"\"") + "\")" : null);
            if (XNode.DeepEquals(original, VisioNativeCellMetadata.Snapshot(cell))) continue;
            string address = ScopedAddress(cell, scope);
            XElement? entry = entries.TryGetValue(address, out XElement? found) && found != null &&
                XNode.DeepEquals(original, found.Element(Metadata + "Snapshot")) ? found : null;
            if (entry == null) {
                added ??= new XElement(Metadata + "NativeCellValues");
                entry = new XElement(Metadata + "Cell", new XAttribute("Address", address));
                added.Add(entry);
            }
            entry.Element(Metadata + "LegacyFont")?.Remove();
            entry.Add(new XElement(Metadata + "LegacyFont", original.Attributes().Select(attribute => new XAttribute(attribute))));
            entry.Element(Metadata + "Snapshot")?.Remove();
            entry.Add(VisioNativeCellMetadata.Snapshot(cell));
        }
        if (added?.HasElements == true) documentRoot.Add(added);
    }

    /// <summary>Rewrites only a numeric constant or GUARD(constant), leaving other formulas unevaluated.</summary>
    internal static string RewriteConstantFormula(string formula, Func<string, string?> map) {
        string value = formula.Trim();
        bool guarded = value.StartsWith("GUARD(", StringComparison.OrdinalIgnoreCase) && value.EndsWith(")", StringComparison.Ordinal);
        string constant = guarded ? value.Substring(6, value.Length - 7).Trim() : value;
        if (!int.TryParse(constant, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id)) return formula;
        string? replacement = map(id.ToString(CultureInfo.InvariantCulture));
        return replacement == null ? formula : guarded ? "GUARD(" + replacement + ")" : replacement;
    }
}
