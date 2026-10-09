using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

/// <summary>A detached source font table shared by a master and its modeled descendants.</summary>
internal sealed class VisioMasterFontScope {
    internal XElement FaceNames { get; }
    internal XElement? Fonts { get; private set; }
    private readonly Dictionary<string, string> _names = new(StringComparer.Ordinal);
    private int _capturedFaces;
    private int _capturedFonts;

    internal VisioMasterFontScope(IEnumerable<XAttribute> attributes, IEnumerable<XElement> faces, XElement? fonts) {
        XNamespace ns = VisioDocument.VisioNamespace;
        FaceNames = new XElement(ns + "FaceNames", attributes.Select(a => new XAttribute(a)),
            faces.Select(face => new XElement(face)));
        Fonts = fonts == null ? null : new XElement(fonts);
        _capturedFaces = FaceNames.Elements().Count();
        _capturedFonts = Fonts?.Elements().Count() ?? 0;
        foreach (XElement face in FaceNames.Elements(ns + "FaceName")) {
            string? id = (string?)face.Attribute("ID"), name = (string?)face.Attribute("Name");
            if (id == null || string.IsNullOrWhiteSpace(name)) continue;
            id = VisioFontWireCodec.Identity(id);
            if (!_names.ContainsKey(id)) _names.Add(id, name!);
        }
    }

    internal string? Name(string id) => _names.TryGetValue(VisioFontWireCodec.Identity(id), out string? name) ? name : null;

    // Native-name decoding can append auxiliary aliases while subsequent masters load.
    // Extend one detached table rather than clone the complete document table per master.
    internal void AddSourceEntries(IEnumerable<XElement> faces, XElement? fonts) {
        foreach (XElement face in faces.Skip(_capturedFaces)) {
            FaceNames.Add(new XElement(face));
            _capturedFaces++;
            string? id = (string?)face.Attribute("ID"), name = (string?)face.Attribute("Name");
            if (id == null || string.IsNullOrWhiteSpace(name)) continue;
            id = VisioFontWireCodec.Identity(id);
            if (!_names.ContainsKey(id)) _names.Add(id, name!);
        }
        if (fonts == null) return;
        Fonts ??= new XElement(fonts.Name, fonts.Attributes().Select(attribute => new XAttribute(attribute)));
        foreach (XElement entry in fonts.Elements().Skip(_capturedFonts)) {
            Fonts.Add(new XElement(entry));
            _capturedFonts++;
        }
    }
}

public partial class VisioDocument {
    private VisioMasterFontScope? _ownedMasterFontScope;
    private readonly Dictionary<VisioMasterFontScope, Dictionary<string, string>> _masterFontOutputIds = new();

    /// <summary>Captures font ownership once, without retaining or modifying the source document.</summary>
    private void CaptureMasterFontScope(VisioMaster master) {
        // Registration in another document must not make later authored children
        // inherit the loaded blueprint's source identity table.
        if (master.NativeFontScope != null) return;
        XElement? fonts = PreservedDocumentElements.FirstOrDefault(element => element.Name == XName.Get("Fonts", VisioNamespace));
        if (master.NativeFontScope == null && master.Shape.NativeFontScope == null) {
            _ownedMasterFontScope ??= new VisioMasterFontScope(PreservedFaceNamesAttributes, PreservedFaceNamesElements, fonts);
            _ownedMasterFontScope.AddSourceEntries(PreservedFaceNamesElements, fonts);
        }
        master.NativeFontScope ??= master.Shape.NativeFontScope ?? _ownedMasterFontScope;
        foreach (VisioShape shape in ShapeTree(master.Shape)) shape.NativeFontScope ??= master.NativeFontScope;
    }

    /// <summary>Imports source tables into this destination, leaving shared master cells and caches intact.</summary>
    private void ImportRegisteredMasterFonts(IEnumerable<VisioPage> pagesToSave) {
        _masterFontOutputIds.Clear();
        VisioPage[] pages = pagesToSave.ToArray();
        var masters = _registeredMasters.Concat(pages.SelectMany(page => BuildEffectiveShapeMasterMap(page).Values)).Distinct();
        var scopes = new HashSet<VisioMasterFontScope>();
        foreach (VisioMaster master in masters) {
            if (master.NativeFontScope != null && scopes.Add(master.NativeFontScope)) ImportMasterFontScope(master.NativeFontScope);
            foreach (VisioShape shape in ShapeTree(master.Shape)) PrepareScopedFontFamily(shape);
        }
        foreach (VisioShape shape in pages.SelectMany(page => page.AllShapes())) {
            if (shape.NativeFontScope != null && scopes.Add(shape.NativeFontScope)) ImportMasterFontScope(shape.NativeFontScope);
            PrepareScopedFontFamily(shape);
        }
    }

    private Dictionary<string, string> ImportMasterFontScope(VisioMasterFontScope scope) {
        if (_masterFontOutputIds.TryGetValue(scope, out Dictionary<string, string>? existing)) return existing;
        Dictionary<string, string> ids = ImportFaceNames(scope.FaceNames);
        _masterFontOutputIds.Add(scope, ids);
        if (scope.Fonts == null) return ids;
        XNamespace ns = VisioNamespace;
        XElement? target = PreservedDocumentElements.FirstOrDefault(element => element.Name == ns + "Fonts");
        if (target == null) {
            target = new XElement(ns + "Fonts", scope.Fonts.Attributes().Select(a => new XAttribute(a)));
            PreservedDocumentElements.Add(target);
        }
        foreach (XElement entry in scope.Fonts.Elements()) {
            var imported = new XElement(entry);
            if ((string?)entry.Attribute("ID") is string source && ids.TryGetValue(VisioFontWireCodec.Identity(source), out string? mapped))
                imported.SetAttributeValue("ID", mapped);
            if (!target.Elements().Any(existing => XNode.DeepEquals(existing, imported))) target.Add(imported);
        }
        return ids;
    }

    private void PrepareScopedFontFamily(VisioShape shape) {
        if (shape.NativeFontScope != null && shape.TextStyle?.FontFamilyAssigned == true &&
            !string.IsNullOrWhiteSpace(shape.TextStyle.FontFamily)) EnsureMasterFontFamily(shape.TextStyle.FontFamily!);
    }

    private string EnsureMasterFontFamily(string family) {
        family = family.Trim();
        XElement? existing = PreservedFaceNamesElements.FirstOrDefault(face => face.Name.LocalName == "FaceName" &&
            string.Equals((string?)face.Attribute("Name"), family, StringComparison.OrdinalIgnoreCase));
        if (existing?.Attribute("ID") is XAttribute id) return id.Value;
        var used = new HashSet<int>(PreservedFaceNamesElements.Attributes("ID").Select(a =>
            int.TryParse(a.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : -1));
        string target = NextFaceNameId(used).ToString(CultureInfo.InvariantCulture);
        PreservedFaceNamesElements.Add(new XElement(XName.Get("FaceName", VisioNamespace),
            new XAttribute("ID", target), new XAttribute("Name", family),
            new XAttribute("UnicodeRanges", "0-255"), new XAttribute("CharSets", "0")));
        return target;
    }

    /// <summary>Rebinds finalized private master XML and its matching guarded snapshots to this document.</summary>
    private void RebindMasterFontOutput(VisioMaster master, XElement output, XElement documentRoot, string scope,
        IReadOnlyDictionary<string, XElement?> entries) {
        if (master.NativeFontScope == null) return;
        var models = GetMasterShapeIdentifiers(master);
        var modeledById = ShapeTree(master.Shape).ToDictionary(shape => VisioShapeFormulaReferences.NormalizeSheetId(models[shape.Id]));
        var mappings = new Dictionary<VisioMasterFontScope, Dictionary<string, string>>();
        foreach (XElement cell in VisioFontWireCodec.FontCells(output)) {
            // Raw roots use the source master's scope. An authored modeled child has
            // destination IDs already, including raw content below that new child.
            VisioMasterFontScope? source = master.NativeFontScope;
            foreach (XElement ancestor in cell.Ancestors(XName.Get("Shape", VisioNamespace))) {
                string id = VisioShapeFormulaReferences.NormalizeSheetId((string?)ancestor.Attribute("ID") ?? string.Empty);
                if (!modeledById.TryGetValue(id, out VisioShape? model)) continue;
                source = model.NativeFontScope;
                break;
            }
            if (source == null) continue;
            if (!mappings.TryGetValue(source, out Dictionary<string, string>? ids)) mappings.Add(source, ids = ImportMasterFontScope(source));
            RebindFontCells(new[] { cell }, ids, entries, scope);
        }
        if (output.Name.LocalName != "MasterContents") return;
        foreach (VisioShape shape in ShapeTree(master.Shape)) {
            string modelId = VisioShapeFormulaReferences.NormalizeSheetId(models[shape.Id]);
            XElement? emitted = output.Descendants(XName.Get("Shape", VisioNamespace)).FirstOrDefault(element =>
                VisioShapeFormulaReferences.NormalizeSheetId((string?)element.Attribute("ID") ?? string.Empty) == modelId);
            if (emitted != null) ReplaceScopedAssignedFont(shape, emitted, scope, entries);
        }
    }

    /// <summary>Rebinds only a copied instance's own cells; nested instances can have different source scopes.</summary>
    private void RebindShapeFontOutput(VisioShape shape, XElement output, XElement documentRoot, string scope,
        IReadOnlyDictionary<string, XElement?> entries) {
        if (shape.NativeFontScope == null) return;
        var ids = ImportMasterFontScope(shape.NativeFontScope);
        RebindFontCells(VisioFontWireCodec.FontCells(output).Where(cell =>
            ReferenceEquals(cell.Ancestors(XName.Get("Shape", VisioNamespace)).FirstOrDefault(), output)), ids, entries, scope);
        ReplaceScopedAssignedFont(shape, output, scope, entries);
    }

    private static void RebindFontCells(IEnumerable<XElement> cells, IReadOnlyDictionary<string, string> ids,
        IReadOnlyDictionary<string, XElement?> entries, string scope) {
        XNamespace metadata = VisioNativeCellMetadata.Namespace;
        foreach (XElement cell in cells) {
            string address = scope + "/" + VisioNativeCellMetadata.Address(cell);
            XElement? entry = entries.TryGetValue(address, out XElement? found) && found != null && VisioNativeCellMetadata.Matches(found, cell) ? found : null;
            RemapImportedFontCell(cell, ids);
            // Only mechanical identity rebinding can preserve producer error/condition provenance.
            if (entry != null) {
                entry.Element(metadata + "Snapshot")!.ReplaceWith(VisioNativeCellMetadata.Snapshot(cell));
                if (entry.Element(metadata + "LegacyFont") is XElement original) RemapImportedFontCell(original, ids);
            }
        }
    }

    private void ReplaceScopedAssignedFont(VisioShape shape, XElement output, string scope,
        IReadOnlyDictionary<string, XElement?> entries) {
        if (shape.TextStyle?.FontFamilyAssigned != true) return;
        XNamespace ns = VisioNamespace;
        XElement? section = output.Elements(ns + "Section").FirstOrDefault(element => (string?)element.Attribute("N") == "Character");
        XElement[] rows = section?.Elements(ns + "Row").ToArray() ?? Array.Empty<XElement>();
        if (rows.Length != 1) return;
        XElement? old = rows[0].Elements(ns + "Cell").FirstOrDefault(element => (string?)element.Attribute("N") == "Font");
        string address = scope + "/" + (old == null ? string.Empty : VisioNativeCellMetadata.Address(old));
        if (entries.TryGetValue(address, out XElement? entry) && entry != null) entry.Remove();
        if (string.IsNullOrWhiteSpace(shape.TextStyle.FontFamily)) { old?.Remove(); return; }
        string id = EnsureMasterFontFamily(shape.TextStyle.FontFamily!);
        var replacement = new XElement(ns + "Cell", new XAttribute("N", "Font"), new XAttribute("V", id));
        if (old == null) rows[0].Add(replacement); else old.ReplaceWith(replacement);
    }

    /// <summary>Returns a name-valued render view, keeping numeric source rows and sentinel cells unchanged.</summary>
    internal static XElement? GetMasterFontRenderSection(VisioShape shape, XElement? section) {
        if (section == null || shape.NativeFontScope == null) return section;
        var output = new XElement(section);
        foreach (XElement cell in VisioFontWireCodec.FontCells(output)) {
            string? value = (string?)cell.Attribute("V");
            if (value != null && !VisioFontWireCodec.IsFallbackSentinel(cell, value) && shape.NativeFontScope.Name(value) is string name)
                cell.SetAttributeValue("V", name);
        }
        if (shape.TextStyle?.FontFamilyAssigned == true && (string?)output.Attribute("N") == "Character") {
            XElement[] rows = output.Elements(XName.Get("Row", VisioNamespace)).ToArray();
            if (rows.Length == 1) {
                XElement? font = rows[0].Elements(XName.Get("Cell", VisioNamespace)).FirstOrDefault(cell => (string?)cell.Attribute("N") == "Font");
                if (string.IsNullOrWhiteSpace(shape.TextStyle.FontFamily)) font?.Remove();
                else {
                    var replacement = new XElement(XName.Get("Cell", VisioNamespace), new XAttribute("N", "Font"), new XAttribute("V", shape.TextStyle.FontFamily!.Trim()));
                    if (font == null) rows[0].Add(replacement); else font.ReplaceWith(replacement);
                }
            }
        }
        return output;
    }
}
