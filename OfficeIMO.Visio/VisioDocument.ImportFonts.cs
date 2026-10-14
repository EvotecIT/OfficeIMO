using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    private Dictionary<string, string> ImportFaceNames(XElement? faceNames) {
        var ids = new Dictionary<string, string>(StringComparer.Ordinal);
        if (faceNames == null) return ids;
        foreach (XAttribute attribute in faceNames.Attributes().Where(ShouldPreserveFaceNamesAttribute))
            AddMissingAttribute(PreservedFaceNamesAttributes, attribute);
        var used = new HashSet<int>();
        foreach (XElement face in PreservedFaceNamesElements)
            if (int.TryParse((string?)face.Attribute("ID"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int id)) used.Add(id);
        foreach (XElement source in faceNames.Elements().Where(ShouldPreserveFaceNamesElement)) {
            if (source.Name.LocalName != "FaceName" || source.Attribute("ID") == null) {
                if (!PreservedFaceNamesElements.Any(element => XNode.DeepEquals(element, source))) PreservedFaceNamesElements.Add(new XElement(source));
                continue;
            }
            string sourceId = (string)source.Attribute("ID")!;
            bool positive = int.TryParse(sourceId, NumberStyles.Integer, CultureInfo.InvariantCulture, out int sourceNumber) && sourceNumber > 0;
            var attributes = source.Attributes().Where(a => !a.IsNamespaceDeclaration && a.Name != "ID").ToArray();
            XElement? existing = PreservedFaceNamesElements.FirstOrDefault(face => face.Name == source.Name &&
                (!positive || VisioFontWireCodec.Identity((string?)face.Attribute("ID") ?? string.Empty) != "0") &&
                face.Attributes().Count(a => !a.IsNamespaceDeclaration && a.Name != "ID") == attributes.Length &&
                attributes.All(a => (string?)face.Attribute(a.Name) == a.Value));
            if (existing?.Attribute("ID") is XAttribute matched) { ids[VisioFontWireCodec.Identity(sourceId)] = matched.Value; continue; }
            var reserved = new HashSet<int>(used);
            if (positive) reserved.Add(0);
            int target = int.TryParse(sourceId, NumberStyles.Integer, CultureInfo.InvariantCulture, out int suggested) && suggested >= 0 && !used.Contains(suggested)
                ? suggested : NextFaceNameId(reserved);
            used.Add(target);
            var imported = new XElement(source);
            string targetId = target.ToString(CultureInfo.InvariantCulture);
            imported.SetAttributeValue("ID", targetId); ids[VisioFontWireCodec.Identity(sourceId)] = targetId;
            PreservedFaceNamesElements.Add(imported);
        }
        return ids;
    }

    private static void RemapImportedFontIds(XElement root, IReadOnlyDictionary<string, string> ids,
        IReadOnlyDictionary<string, XElement?>? nativeCells = null, string? nativeScope = null) {
        foreach (XElement cell in VisioFontWireCodec.FontCells(root)) {
            string address = (nativeScope == null ? string.Empty : nativeScope + "/") + VisioNativeCellMetadata.Address(cell);
            XElement? snapshot = nativeCells != null && nativeCells.TryGetValue(address, out XElement? entry) && entry != null && VisioNativeCellMetadata.Matches(entry, cell)
                ? entry.Element(VisioNativeCellMetadata.Namespace + "Snapshot") : null;
            RemapImportedFontCell(cell, ids);
            // Font rebinding changes identity, not the producer's cached result. Rebase
            // only a snapshot that matched before the mechanical rewrite.
            if (snapshot != null) RemapImportedFontCell(snapshot, ids);
        }
    }

    private static void RemapImportedFontCell(XElement cell, IReadOnlyDictionary<string, string> ids) {
        if (VisioFontWireCodec.IsFallbackSentinel(cell, (string?)cell.Attribute("V"))) return;
        if ((string?)cell.Attribute("V") is string id && ids.TryGetValue(VisioFontWireCodec.Identity(id), out string? replacement)) cell.SetAttributeValue("V", replacement);
        if (cell.Attribute("F") is not XAttribute formula) return;
        formula.Value = VisioFontWireCodec.RewriteConstantFormula(formula.Value, id => ids.TryGetValue(id, out string? mapped) ? mapped : null);
    }
}
