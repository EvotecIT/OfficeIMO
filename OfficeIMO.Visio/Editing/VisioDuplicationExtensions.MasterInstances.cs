using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

public static partial class VisioDuplicationExtensions {
    // Uses the same deep-copy owner as page/selection duplication. Master definitions and
    // sibling instances remain independent; the master link supplies inherited geometry.
    internal static VisioShape CreateMasterInstance(VisioPage page, VisioMaster master,
        string id, double x, double y, double width, double height, string? text) {
        if (page.AllShapes().Any(shape => shape.Id == id) || page.Connectors.Any(connector => connector.Id == id))
            throw new ArgumentException("A shape or connector with this identifier already exists on the page.", nameof(id));
        var map = new Dictionary<VisioShape, VisioShape>();
        var root = CloneShape(master.Shape, new IdAllocator(page), 0, 0, false, map,
            new Dictionary<VisioConnectionPoint, VisioConnectionPoint>(),
            new VisioShapeDuplicationOptions { ShapeIdFactory = source => ReferenceEquals(source, master.Shape) ? id : id + ":" + source.Id });
        var masterIds = VisioDocument.GetMasterShapeIdentifiers(master);
        VisioDocument.ValidateCopiedNativeCellMetadata(map.Values, Array.Empty<VisioConnector>(), page.OwnerDocument, page);
        // Resolve the same simple-master frame that saving emits before resizing it.
        // This also covers a blank blueprint whose text override is applied below.
        VisioTextStyle? simpleFrame = VisioDocument.CreateSimpleMasterTextFrameStyle(master, master.Shape.WidthInches, master.Shape.HeightInches);
        if (simpleFrame != null) root.TextStyle = simpleFrame;
        double scaleX = master.Shape.WidthInches > 0 ? width / master.Shape.WidthInches : 1;
        double scaleY = master.Shape.HeightInches > 0 ? height / master.Shape.HeightInches : 1;
        foreach (var pair in map) {
            VisioShape source = pair.Key, instance = pair.Value;
            instance.Master = master;
            instance.MasterShape = source;
            instance.MasterShapeId = ReferenceEquals(instance, root) ? null : masterIds[source.Id];
            instance.NameU ??= master.NameU;
            instance.Type = source.Children.Count > 0 ? "Group" : source.Type;
            // Geometry formulas must see the instance's final placement when
            // deciding whether a formula still matches its resized cache.
            instance.PinXInches = ReferenceEquals(instance, root) ? x : instance.PinXInches * scaleX;
            instance.PinYInches = ReferenceEquals(instance, root) ? y : instance.PinYInches * scaleY;
            instance.WidthInches *= scaleX; instance.HeightInches *= scaleY;
            instance.LocPinXInches *= scaleX; instance.LocPinYInches *= scaleY;
            instance.HasExplicitLocPinX = true; instance.HasExplicitLocPinY = true;
            instance.TextStyle?.ScaleTextBlock(scaleX, scaleY);
            if (!string.IsNullOrEmpty(instance.Text)) {
                instance.TextStyle ??= new VisioTextStyle();
                instance.TextStyle.TextWidth ??= instance.WidthInches;
                instance.TextStyle.TextHeight ??= instance.HeightInches;
                instance.TextStyle.TextPinX ??= instance.WidthInches / 2;
                instance.TextStyle.TextPinY ??= instance.HeightInches / 2;
                instance.TextStyle.TextLocPinX ??= instance.TextStyle.TextWidth / 2;
                instance.TextStyle.TextLocPinY ??= instance.TextStyle.TextHeight / 2;
            }
            foreach (VisioConnectionPoint point in instance.ConnectionPoints) { point.X *= scaleX; point.Y *= scaleY; }
            VisioGeometryScaling.Scale(instance, scaleX, scaleY, source);
            if (instance.PreservedShapeChildren.Count == 0) {
                foreach (var entry in source.PreservedShapeChildren)
                    instance.PreservedShapeChildren.Add(entry.RawElement != null
                        ? new VisioShape.PreservedShapeChildEntry(entry.RawElement)
                        : new VisioShape.PreservedShapeChildEntry(entry.Token!));
            }
        }
        root.PinXInches = x; root.PinYInches = y; root.WidthInches = width; root.HeightInches = height;
        root.NameU = master.NameU;
        if (text != null) root.Text = text;
        RemapContainerMembership(map);
        page.Shapes.Add(root);
        try {
            var ids = VisioDocument.AssignPageElementIdentifiers(page);
            var references = BuildFormulaReferenceMap(map.Select(pair => (masterIds[pair.Key.Id], ids[pair.Value.Id])));
            foreach (VisioShape instance in map.Values) VisioShapeFormulaReferences.Remap(instance, references);
            // MasterShape links must remain stable when callers later insert/reorder
            // children in a generated master that did not have native IDs yet.
            foreach (VisioShape source in map.Keys) source.PersistedId = masterIds[source.Id];
            return root;
        } catch {
            page.Shapes.Remove(root);
            throw;
        }
    }
}
