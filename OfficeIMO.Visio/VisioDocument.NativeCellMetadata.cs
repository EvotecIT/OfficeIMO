using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    internal static void ValidateCopiedNativePageSheetMetadata(VisioPage page) {
        if (page.NativePageSheetMetadata == null) return;
        var xml = new XDocument();
        using (var writer = xml.CreateWriter()) new VisioDocument().WritePageSheet(writer, VisioNamespace, page);
        page.NativePageSheetMetadata = page.NativePageSheetMetadata.Validated(xml.Root!, LayerValueAssignments(page));
    }

    private static ISet<string> LayerValueAssignments(VisioPage page) {
        BuildLayerIndexMap(page, out List<VisioLayer> layers);
        return VisioNativeCellAssignments.ForLayers(layers);
    }

    /// <summary>Checks copied provenance with the canonical shape writer, before mechanical formula rebinding.</summary>
    internal static void ValidateCopiedNativeCellMetadata(IEnumerable<VisioShape> shapes, IEnumerable<VisioConnector> connectors, VisioDocument? destination = null, VisioPage? page = null) {
        var copied = new HashSet<VisioShape>(shapes);
        var layerIndexes = page != null ? BuildLayerIndexMap(page, out _) : new Dictionary<string, int>();
        var writerOwner = new VisioDocument { UseMastersByDefault = false };
        if (destination != null)
            foreach (XElement color in destination.PreservedColorsElements) writerOwner.PreservedColorsElements.Add(new XElement(color));
        foreach (VisioShape root in copied.Where(s => s.Parent == null || !copied.Contains(s.Parent))) {
            if (!ShapeTree(root).Any(s => s.NativeCellMetadata != null)) continue;
            ValidateTree(root, writerOwner.CreateMasterModelShapeXml(root, layerIndexes: layerIndexes), writerOwner);
        }
        foreach (VisioConnector connector in connectors.Where(c => c.NativeCellMetadata != null)) {
            var xml = new XDocument();
            var ids = new Dictionary<string, string>(StringComparer.Ordinal) { [connector.Id] = "1" };
            if (connector.From != null) ids[connector.From.Id] = connector.From.PersistedId ?? connector.From.Id;
            if (connector.To != null) ids[connector.To.Id] = connector.To.PersistedId ?? connector.To.Id;
            using (var writer = xml.CreateWriter()) {
                writer.WriteStartElement("Shape", VisioNamespace);
                writerOwner.WriteConnectorShapeBody(writer, VisioNamespace, connector, ids, layerIndexes);
                writer.WriteEndElement();
            }
            connector.NativeCellMetadata = connector.NativeCellMetadata!.Validated(xml.Root!, VisioNativeCellAssignments.For(connector, writerOwner));
        }
    }

    private static void ValidateTree(VisioShape model, XElement xml, VisioDocument writerOwner) {
        model.NativeCellMetadata = model.NativeCellMetadata?.Validated(xml, VisioNativeCellAssignments.For(model, writerOwner));
        XElement[] children = xml.Element(XName.Get("Shapes", VisioNamespace))?.Elements(XName.Get("Shape", VisioNamespace)).ToArray() ?? Array.Empty<XElement>();
        for (int i = 0; i < model.Children.Count; i++) {
            if (i < children.Length) ValidateTree(model.Children[i], children[i], writerOwner);
            else foreach (VisioShape child in ShapeTree(model.Children[i])) child.NativeCellMetadata = null;
        }
    }

    // Collection reads only relevant already-written parts, once per save. It never serializes
    // a source document/package per copy or keeps a reference to the source document.
    private void CollectNativeCellMetadata(XDocument documentXml, Package package,
        IEnumerable<(VisioPage Page, PackagePart Part, PackageRelationship Relationship)> pages, IReadOnlyList<PackageMasterEntry> masters,
        PackagePart pagesPart, PackagePart? mastersPart) {
        var metadata = new XElement(VisioNativeCellMetadata.Namespace + "NativeCellValues");
        foreach (var page in pages) {
            var shapes = page.Page.AllShapes().Where(s => s.NativeCellMetadata != null).ToArray();
            var connectors = page.Page.Connectors.Where(c => c.NativeCellMetadata != null).ToArray();
            if (shapes.Length == 0 && connectors.Length == 0) continue;
            var xml = LoadPackageXml(page.Part, "Written Visio page cell metadata");
            var cells = ShapeElementsById(xml);
            var ids = AssignPageElementIdentifiers(page.Page);
            string scope = VisioNativeCellMetadata.PageScope(page.Page.Id);
            foreach (VisioShape shape in shapes)
                if (cells.TryGetValue(VisioShapeFormulaReferences.NormalizeSheetId(ids[shape.Id]), out XElement? emitted) && emitted != null) shape.NativeCellMetadata!.Collect(emitted, scope, metadata, VisioNativeCellAssignments.For(shape, this));
            foreach (VisioConnector connector in connectors)
                if (cells.TryGetValue(VisioShapeFormulaReferences.NormalizeSheetId(ids[connector.Id]), out XElement? emitted) && emitted != null) connector.NativeCellMetadata!.Collect(emitted, scope, metadata, VisioNativeCellAssignments.For(connector, this));
        }
        foreach (PackageMasterEntry master in masters) {
            var shapes = ShapeTree(master.Master.Shape).Where(s => s.NativeCellMetadata != null).ToArray();
            if (shapes.Length == 0 && !master.Master.NativeAdditionalShapeMetadata.Values.Any(value => value != null)) continue;
            var part = package.GetPart(new Uri("/visio/masters/master" + master.PartNumber + ".xml", UriKind.Relative));
            var xml = LoadPackageXml(part, "Written Visio master cell metadata");
            var cells = ShapeElementsById(xml);
            var ids = GetMasterShapeIdentifiers(master.Master);
            string scope = VisioNativeCellMetadata.MasterScope(master.Master.NameU);
            foreach (VisioShape shape in shapes)
                if (cells.TryGetValue(VisioShapeFormulaReferences.NormalizeSheetId(ids[shape.Id]), out XElement? emitted) && emitted != null) shape.NativeCellMetadata!.Collect(emitted, scope, metadata, VisioNativeCellAssignments.For(shape));
            foreach (var raw in master.Master.NativeAdditionalShapeMetadata)
                if (raw.Value != null && cells.TryGetValue(raw.Key, out XElement? emitted) && emitted != null)
                    raw.Value.Collect(emitted, scope, metadata);
        }
        var pageSheets = pages.Select(p => p.Page).Where(p => p.NativePageSheetMetadata != null).ToArray();
        if (pageSheets.Length > 0) {
            var xml = LoadPackageXml(pagesPart, "Written Visio page catalog cell metadata");
            foreach (VisioPage page in pageSheets) {
                XElement? sheet = xml.Root?.Elements(XName.Get("Page", VisioNamespace)).SingleOrDefault(p => (string?)p.Attribute("ID") == page.Id.ToString(System.Globalization.CultureInfo.InvariantCulture))?.Element(XName.Get("PageSheet", VisioNamespace));
                if (sheet != null) page.NativePageSheetMetadata!.Collect(sheet, "Pages", metadata, LayerValueAssignments(page));
            }
        }
        var masterSheets = masters.Where(m => m.Master.NativePageSheetMetadata != null).ToArray();
        if (masterSheets.Length > 0 && mastersPart != null) {
            var xml = LoadPackageXml(mastersPart, "Written Visio master catalog cell metadata");
            foreach (PackageMasterEntry master in masterSheets) {
                XElement? sheet = xml.Root?.Elements(XName.Get("Master", VisioNamespace)).SingleOrDefault(m => (string?)m.Attribute("ID") == master.PackageId)?.Element(XName.Get("PageSheet", VisioNamespace));
                if (sheet != null) master.Master.NativePageSheetMetadata!.Collect(sheet, "Masters", metadata);
            }
        }
        if (metadata.HasElements) documentXml.Root!.Add(metadata);
        foreach (XElement empty in documentXml.Root!.Elements(VisioNativeCellMetadata.Namespace + "NativeCellValues").Where(e => !e.HasElements).ToArray()) empty.Remove();
    }

    private static IEnumerable<VisioShape> ShapeTree(VisioShape shape) {
        yield return shape;
        foreach (VisioShape child in shape.Children)
            foreach (VisioShape descendant in ShapeTree(child)) yield return descendant;
    }

    private static Dictionary<string, XElement?> ShapeElementsById(XDocument xml) {
        var elements = new Dictionary<string, XElement?>(StringComparer.Ordinal);
        foreach (XElement shape in xml.Descendants(XName.Get("Shape", VisioNamespace))) {
            string? id = (string?)shape.Attribute("ID");
            if (id == null) continue;
            id = VisioShapeFormulaReferences.NormalizeSheetId(id);
            if (elements.ContainsKey(id)) elements[id] = null;
            else elements.Add(id, shape);
        }
        return elements;
    }
}
