using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

/// <summary>
/// Translates native name-valued fonts and legacy numeric identities at the XML boundary.
/// The shared model keeps legacy identities, including distinct charset entries for one family.
/// </summary>
internal sealed partial class VisioFontWireCodec {
    private static readonly XNamespace Modern = VisioDocument.VisioNamespace;
    private static readonly XNamespace Metadata = VisioNativeCellMetadata.Namespace;
    private static readonly string[] FaceAttributes = { "CharSets", "Flags", "Panos", "UnicodeRanges" };
    private readonly Dictionary<string, string> _namesById = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _idsByName = new(StringComparer.OrdinalIgnoreCase);
    private readonly XElement? _faces;
    private readonly bool _restoredLegacyFonts;
    private IList<XElement>? _modelFaces;
    private Action<XElement>? _aliasImporter;

    private VisioFontWireCodec(XElement? faces, bool restoredLegacyFonts = false) {
        _faces = faces;
        _restoredLegacyFonts = restoredLegacyFonts;
        foreach (XElement face in faces?.Elements(Modern + "FaceName") ?? Enumerable.Empty<XElement>()) {
            string? id = (string?)face.Attribute("ID"), name = (string?)face.Attribute("Name");
            if (id == null || string.IsNullOrWhiteSpace(name)) continue;
            id = Identity(id);
            if (!_namesById.ContainsKey(id)) _namesById.Add(id, name!);
            if (!_idsByName.ContainsKey(name!)) _idsByName.Add(name!, id);
        }
    }

    /// <summary>Shares only the detached font table with the loader, so auxiliary aliases remain resolvable.</summary>
    internal void BindModelFaces(IList<XElement> faces) => _modelFaces = faces;

    /// <summary>Transfers aliases created while later package parts are decoded into the destination map.</summary>
    internal void BindAliasImporter(Action<XElement> importer) => _aliasImporter = importer;

    internal static string Identity(string value) => int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id)
        ? id.ToString(CultureInfo.InvariantCulture) : value;

    /// <summary>Normalizes a private input tree before model baselines and native snapshots are bound.</summary>
    internal static VisioFontWireCodec ReadDocument(XElement root) {
        XElement? faces = root.Element(Modern + "FaceNames");
        XElement[] saved = root.Elements(Metadata + "FontTable").ToArray();
        bool restored = false;
        if (saved.Length == 1 && XNode.DeepEquals(faces, UnpackTable(saved[0].Element(Metadata + "Snapshot")?.Element(Metadata + "FaceNames")))) {
            restored = true;
            XElement? original = UnpackTable(saved[0].Element(Metadata + "Original")?.Element(Metadata + "FaceNames"));
            faces?.Remove();
            faces = original == null ? null : new XElement(original);
            if (faces != null) root.Add(faces);
            XElement? fonts = UnpackTable(saved[0].Element(Metadata + "Original")?.Element(Metadata + "Fonts"));
            root.Element(Modern + "Fonts")?.Remove();
            if (fonts != null) root.Add(new XElement(fonts));
        }
        foreach (XElement container in saved) container.Remove();
        // Invalidate the payload itself across all part scopes, not just its use on
        // this load. A literal family may be redeclared on save without changing
        // the cell, and must not resurrect an obsolete ID on the following reopen.
        if (!restored)
            foreach (XElement original in root.Elements(Metadata + "NativeCellValues").Elements(Metadata + "Cell")
                         .Elements(Metadata + "LegacyFont").ToArray()) original.Remove();
        NormalizeFaceTable(faces);
        return new VisioFontWireCodec(faces, restored);
    }

    // Keep preservation payloads outside the native element namespace. Consumers searching
    // for native FaceName nodes must see the actual table, not copies of legacy evidence.
    private static XElement? UnpackTable(XElement? table) => TranslateTable(table, Metadata, Modern);
    private static XElement? PackTable(XElement? table) => TranslateTable(table, Modern, Metadata);
    private static XElement? TranslateTable(XElement? table, XNamespace from, XNamespace to) {
        if (table == null) return null;
        var copy = new XElement(table);
        foreach (XElement element in copy.DescendantsAndSelf()) if (element.Name.Namespace == from) element.Name = to + element.Name.LocalName;
        return copy;
    }

    private static void NormalizeFaceTable(XElement? faces) {
        if (faces == null) return;
        var used = new HashSet<int>(faces.Elements(Modern + "FaceName").Attributes("ID")
            .Select(a => int.TryParse(a.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id) ? id : -1));
        // Zero is an auxiliary-font fallback sentinel. Native names receive positive
        // internal identities so Asian/complex/bullet fonts cannot turn into inheritance.
        int next = 1;
        foreach (XElement face in faces.Elements(Modern + "FaceName")) {
            string? name = (string?)face.Attribute("NameU") ?? (string?)face.Attribute("Name");
            if (name == null) continue;
            if (face.Attribute("ID") == null) {
                while (used.Contains(next)) next++;
                face.SetAttributeValue("ID", next); used.Add(next++);
            }
            face.SetAttributeValue("Name", name);
            face.Attribute("NameU")?.Remove();
        }
    }

    /// <summary>Enumerates only the four ShapeSheet font cells, preserving fallback sentinels.</summary>
    internal static IEnumerable<XElement> FontCells(XElement root) => root.DescendantsAndSelf(Modern + "Cell").Where(cell => {
        string? section = (string?)cell.Parent?.Parent?.Attribute("N"), name = (string?)cell.Attribute("N");
        return section == "Character" && (name is "Font" or "AsianFont" or "ComplexScriptFont") ||
               section == "Paragraph" && name == "BulletFont";
    });

    internal static bool IsFallbackSentinel(XElement cell, string? value) =>
        (string?)cell.Attribute("N") != "Font" && (string.IsNullOrEmpty(value) || Identity(value!) == "0");

    /// <summary>Restores guarded numeric cells, or binds standards-shaped font names to internal IDs.</summary>
    internal void DecodeCells(XElement root, IReadOnlyDictionary<string, XElement?> entries, string? scope = null) {
        foreach (XElement cell in FontCells(root)) {
            string address = ScopedAddress(cell, scope);
            XElement? entry = entries.TryGetValue(address, out XElement? found) && found != null && VisioNativeCellMetadata.Matches(found, cell) ? found : null;
            // A numeric identity is meaningful only with the table that declared it.
            // External table edits keep current native names authoritative for every part.
            if (_restoredLegacyFonts && entry?.Element(Metadata + "LegacyFont") is XElement original) {
                ReplaceValues(cell, original);
            } else {
                string? value = (string?)cell.Attribute("V");
                if (!IsFallbackSentinel(cell, value) && value != null && _idsByName.TryGetValue(value, out string? id)) {
                    // An external native edit can select the family whose preserved legacy
                    // ID is zero. Give auxiliary named references a positive alias instead
                    // of silently changing that selection into a fallback sentinel.
                    if ((string?)cell.Attribute("N") != "Font" && id == "0") id = PositiveAlias(value);
                    cell.SetAttributeValue("V", id);
                }
            }
            if (entry != null) entry.Element(Metadata + "Snapshot")!.ReplaceWith(VisioNativeCellMetadata.Snapshot(cell));
        }
    }

    private string PositiveAlias(string family) {
        foreach (var pair in _namesById)
            if (pair.Key != "0" && string.Equals(pair.Value, family, StringComparison.OrdinalIgnoreCase)) return pair.Key;
        int next = 1;
        while (_namesById.ContainsKey(next.ToString(CultureInfo.InvariantCulture))) next++;
        string id = next.ToString(CultureInfo.InvariantCulture);
        var face = new XElement(Modern + "FaceName", new XAttribute("ID", id), new XAttribute("Name", family));
        _faces?.Add(face); _modelFaces?.Add(new XElement(face)); _namesById.Add(id, family);
        _aliasImporter?.Invoke(new XElement(face));
        return id;
    }

    private static void ReplaceValues(XElement cell, XElement snapshot) {
        foreach (string name in new[] { "N", "V", "F", "U", "E" }) cell.SetAttributeValue(name, (string?)snapshot.Attribute(name));
    }

    private static string ScopedAddress(XElement cell, string? scope) =>
        (scope == null ? string.Empty : scope + "/") + VisioNativeCellMetadata.Address(cell);
}
