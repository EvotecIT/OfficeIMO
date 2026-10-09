using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

public partial class VisioDocument {
    internal static Dictionary<string, string> GetMasterShapeIdentifiers(VisioMaster master) =>
        BuildPersistedIdMap(new[] { master.Shape }, Array.Empty<VisioConnector>(), new Dictionary<string, VisioMaster>(),
            master.RawMasterContentXml?.Descendants(XName.Get("Shape", VisioNamespace)).Attributes("ID").Select(attribute => attribute.Value)
            ?? master.PreservedAdditionalShapeElements.SelectMany(element => element.DescendantsAndSelf(XName.Get("Shape", VisioNamespace))).Attributes("ID").Select(attribute => attribute.Value));

    // Formulas refer to native Sheet IDs. Pin those IDs before copying formulas so
    // subsequent insertion/reordering cannot redirect their targets.
    internal static Dictionary<string, string> GetPageElementIdentifiers(VisioPage page, IEnumerable<VisioShape>? addedTrees = null) {
        var masters = new Dictionary<string, VisioMaster>();
        foreach (VisioShape shape in page.AllShapes()) {
            VisioMaster? master = shape.Master;
            if (master == null && page.OwnerDocument?.UseMastersByDefault == true && !string.IsNullOrWhiteSpace(shape.NameU))
                page.OwnerDocument.TryGetMaster(shape.NameU!, out master);
            if (master != null) masters[shape.Id] = master;
        }
        return BuildPersistedIdMap(page.Shapes.Concat(addedTrees ?? Array.Empty<VisioShape>()), page.Connectors, masters);
    }

    internal static Dictionary<string, string> AssignPageElementIdentifiers(VisioPage page) {
        var ids = GetPageElementIdentifiers(page);
        foreach (VisioShape shape in page.AllShapes()) shape.PersistedId = ids[shape.Id];
        foreach (VisioConnector connector in page.Connectors) connector.PersistedId = ids[connector.Id];
        return ids;
    }

    private static Dictionary<string, string> BuildPersistedIdMap(IEnumerable<VisioShape> shapes,
        IEnumerable<VisioConnector> connectors, IReadOnlyDictionary<string, VisioMaster> effectiveMasters,
        IEnumerable<string>? reservedIds = null) {
        var candidates = new Dictionary<string, string?>(StringComparer.Ordinal);
        var used = new HashSet<uint>();
        var assigned = new HashSet<uint>();
        foreach (string id in reservedIds ?? Array.Empty<string>())
            if (uint.TryParse(id, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint numeric)) used.Add(numeric);

        void Visit(VisioShape shape) {
            if (candidates.ContainsKey(shape.Id)) throw new InvalidOperationException("Duplicate shape identifier: " + shape.Id);
            candidates.Add(shape.Id, shape.PersistedId);
            if (effectiveMasters.TryGetValue(shape.Id, out VisioMaster? master) && UsesRawMasterInstanceChildren(shape, master))
                ReserveRawMasterInstanceChildIds(shape, master, id => { if (!candidates.ContainsKey(id)) candidates.Add(id, null); });
            foreach (VisioShape child in shape.Children) Visit(child);
        }
        foreach (VisioShape shape in shapes) Visit(shape);
        foreach (VisioConnector connector in connectors) {
            if (candidates.ContainsKey(connector.Id)) throw new InvalidOperationException("Duplicate shape identifier: " + connector.Id);
            candidates.Add(connector.Id, connector.PersistedId);
        }

        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        // Loaded native IDs have priority over new API IDs, regardless of collection order.
        // They may also occur in reserved source XML; only their original model can reclaim them.
        foreach (var candidate in candidates) {
            if (uint.TryParse(candidate.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint numeric) && assigned.Add(numeric)) {
                used.Add(numeric); map.Add(candidate.Key, candidate.Value!);
            }
        }
        uint next = 1;
        foreach (var candidate in candidates) {
            if (map.ContainsKey(candidate.Key)) continue;
            if (uint.TryParse(candidate.Key, NumberStyles.Integer, CultureInfo.InvariantCulture, out uint numeric) && used.Add(numeric)) {
                map.Add(candidate.Key, candidate.Key);
                continue;
            }
            while (used.Contains(next)) next++;
            used.Add(next); map.Add(candidate.Key, next++.ToString(CultureInfo.InvariantCulture));
        }
        return map;
    }
}
