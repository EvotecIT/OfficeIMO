using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

internal sealed partial class VisioLegacyBinaryCodec {
    private XElement ReadMaster(VisioBinaryContainer.Node node, string name) {
        var master = new XElement(Ns + "Master", Id(node.Id), new XAttribute("Name", name), new XAttribute("NameU", name));
        var shapes = new List<ShapeEntry>();
        foreach (var child in node.Children) {
            if (child.Type == 0x46) master.Add(ReadSheet(Chunks(child), "PageSheet", child.Id));
            else if (child.Type is 0x47 or 0x48 or 0x4d or 0x4e)
                shapes.Add(ReadShape(Chunks(child), child.Id));
        }
        master.Add(AssembleShapes(shapes));
        return master;
    }

    private XElement ReadPage(VisioBinaryContainer.Node node, string name) {
        var chunks = Chunks(node);
        var page = new XElement(Ns + "Page", Id(node.Id), new XAttribute("Name", name), new XAttribute("NameU", name),
            new XAttribute("Background", (node.Format & 1) == 0 ? "1" : "0"));
        var shapes = new List<ShapeEntry>();
        int sheetStart = -1;
        for (int i = 0; i < chunks.Count; i++) {
            var chunk = chunks[i];
            if (chunk.Type == 0x15) {
                uint background = chunk.U32(8);
                if (background != uint.MaxValue) page.SetAttributeValue("BackPage", background);
            } else if (chunk.Type == 0x46) sheetStart = i;
            else if (chunk.Type is 0x47 or 0x48 or 0x4d or 0x4e) {
                if (sheetStart >= 0) { page.Add(ReadSheet(chunks.GetRange(sheetStart, i - sheetStart), "PageSheet", 0)); sheetStart = -1; }
                int end = i + 1;
                while (end < chunks.Count && chunks[end].Type is not (0x47 or 0x48 or 0x4d or 0x4e)) end++;
                shapes.Add(ReadShape(chunks.GetRange(i, end - i), chunk.Id));
                i = end - 1;
            }
        }
        if (sheetStart >= 0) page.Add(ReadSheet(chunks.GetRange(sheetStart, chunks.Count - sheetStart), "PageSheet", 0));
        page.Add(AssembleShapes(shapes));
        return page;
    }

    private ShapeEntry ReadShape(List<VisioBinaryChunks.Chunk> chunks, uint pointerId) {
        var header = chunks.FirstOrDefault(chunk => chunk.Type is 0x47 or 0x48 or 0x4d or 0x4e);
        if (header.Length == 0) throw new InvalidDataException("Binary Visio shape header is missing.");
        uint id = header.Id == uint.MaxValue ? pointerId : header.Id;
        uint parent = header.U32(10), master = header.U32(18), masterShape = header.U32(26);
        Item();
        var shape = ReadSheet(chunks, "Shape", id);
        shape.SetAttributeValue("Type", header.Type == 0x47 ? "Group" : "Shape");
        if (master != uint.MaxValue) shape.SetAttributeValue("Master", master);
        if (masterShape != uint.MaxValue) shape.SetAttributeValue("MasterShape", masterShape);
        ApplyStyleReferences(shape, header);
        return new ShapeEntry(id, parent, shape);
    }

    private XElement AssembleShapes(List<ShapeEntry> entries) {
        var result = new XElement(Ns + "Shapes");
        var byId = new Dictionary<uint, ShapeEntry>();
        foreach (var entry in entries) {
            if (byId.ContainsKey(entry.Id)) throw new InvalidDataException("Binary Visio contains duplicate shape identities.");
            byId.Add(entry.Id, entry);
        }
        foreach (var entry in entries) {
            int depth = 0;
            uint current = entry.Parent;
            var visited = new HashSet<uint> { entry.Id };
            while (current != 0 && current != uint.MaxValue) {
                if (!visited.Add(current)) throw new InvalidDataException("Binary Visio shape hierarchy contains a cycle.");
                if (++depth > _options.MaxDepth) throw new InvalidDataException("Binary Visio shape nesting budget exceeded.");
                if (!byId.TryGetValue(current, out ShapeEntry? parent)) throw new InvalidDataException("Binary Visio shape parent is missing.");
                if ((string?)parent.Xml.Attribute("Type") != "Group") throw new InvalidDataException("Binary Visio shape parent is not a group.");
                current = parent.Parent;
            }
            if (entry.Parent == 0 || entry.Parent == uint.MaxValue) result.Add(entry.Xml);
            else {
                XElement parent = byId[entry.Parent].Xml;
                XElement? children = parent.Element(Ns + "Shapes");
                if (children == null) { children = new XElement(Ns + "Shapes"); parent.Add(children); }
                children.Add(entry.Xml);
            }
        }
        return result;
    }

    private static void ApplyStyleReferences(XElement sheet, VisioBinaryChunks.Chunk header) {
        foreach (var pair in new[] { ("FillStyle", header.Type == 0x4a ? 42 : 34), ("LineStyle", header.Type == 0x4a ? 34 : 42), ("TextStyle", 50) }) {
            uint value = header.U32(pair.Item2);
            if (value != uint.MaxValue) sheet.SetAttributeValue(pair.Item1, value.ToString(CultureInfo.InvariantCulture));
        }
    }

    private sealed class ShapeEntry {
        internal ShapeEntry(uint id, uint parent, XElement xml) { Id = id; Parent = parent; Xml = xml; }
        internal uint Id { get; }
        internal uint Parent { get; }
        internal XElement Xml { get; }
    }
}
