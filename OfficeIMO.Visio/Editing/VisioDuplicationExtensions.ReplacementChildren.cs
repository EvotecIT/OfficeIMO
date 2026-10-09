using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio;

public static partial class VisioDuplicationExtensions {
    /// <summary>Clones a replacement's descendants without touching the page or the retained live root.</summary>
    internal static Dictionary<VisioShape, VisioShape> PrepareReplacementChildren(VisioPage page, VisioShape root, VisioMaster master) {
        var rootReference = new VisioShape(root.Id);
        var map = new Dictionary<VisioShape, VisioShape> { [master.Shape] = rootReference };
        var allocator = new IdAllocator(page);
        var points = new Dictionary<VisioConnectionPoint, VisioConnectionPoint>();
        var options = new VisioShapeDuplicationOptions { ShapeIdFactory = source => root.Id + ":" + source.Id };
        foreach (VisioShape child in master.Shape.Children) CloneShape(child, allocator, 0, 0, false, map, points, options);
        RemapContainerMembership(map);
        map[master.Shape] = root;
        var ids = VisioDocument.GetMasterShapeIdentifiers(master);
        foreach (var pair in map) {
            if (ReferenceEquals(pair.Value, root)) continue;
            VisioShape source = pair.Key, clone = pair.Value;
            clone.Master = master; clone.MasterShape = source; clone.MasterShapeId = ids[source.Id];
            clone.Type = source.Children.Count > 0 ? "Group" : source.Type;
            if (clone.PreservedShapeChildren.Count == 0) foreach (var entry in source.PreservedShapeChildren)
                clone.PreservedShapeChildren.Add(entry.RawElement != null
                    ? new VisioShape.PreservedShapeChildEntry(entry.RawElement)
                    : new VisioShape.PreservedShapeChildEntry(entry.Token!));
        }
        VisioDocument.ValidateCopiedNativeCellMetadata(map.Values.Where(shape => !ReferenceEquals(shape, root)), Array.Empty<VisioConnector>(), page.OwnerDocument, page);
        return map;
    }
}
