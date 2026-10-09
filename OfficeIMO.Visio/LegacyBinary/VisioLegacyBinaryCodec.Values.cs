using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    internal sealed partial class VisioLegacyBinaryCodec {
        private XElement ReadSheet(List<VisioBinaryChunks.Chunk> chunks, string kind, uint id) {
            var sheet = new XElement(Ns + kind, Id(id));
            XElement? geometry = null;
            var names = new Dictionary<uint, string>();
            var fields = new SortedDictionary<uint, string>();
            foreach (var chunk in chunks.Where(chunk => chunk.Type == 0x2d)) names[chunk.Id] = Text(chunk.Text(0));
            foreach (var chunk in chunks) {
                _token.ThrowIfCancellationRequested();
                switch (chunk.Type) {
                    case 0x9b: case 0x9c:
                        string[] transforms = chunk.Type == 0x9b
                            ? new[] { "PinX", "PinY", "Width", "Height", "LocPinX", "LocPinY", "Angle" }
                            : new[] { "TxtPinX", "TxtPinY", "TxtWidth", "TxtHeight", "TxtLocPinX", "TxtLocPinY", "TxtAngle" };
                        var xform = new XElement(Ns + (chunk.Type == 0x9b ? "XForm" : "TextXForm"));
                        for (int index = 0; index < transforms.Length; index++) xform.Add(Cell(transforms[index], chunk.Number(1 + index * 9)));
                        if (chunk.Type == 0x9b) xform.Add(Cell("FlipX", chunk.Byte(63)), Cell("FlipY", chunk.Byte(64)));
                        sheet.Add(xform); break;
                    case 0x9d:
                        sheet.Add(new XElement(Ns + "XForm1D", Cell("BeginX", chunk.Number(1)), Cell("BeginY", chunk.Number(10)),
                            Cell("EndX", chunk.Number(19)), Cell("EndY", chunk.Number(28)))); break;
                    case 0x92:
                        double width = chunk.Number(1), height = chunk.Number(10), pageScale = chunk.Number(37), drawingScale = chunk.Number(46);
                        if (width <= 0 || height <= 0 || pageScale <= 0 || drawingScale <= 0)
                            throw new InvalidDataException("Binary Visio page size and scale must be positive.");
                        sheet.Add(new XElement(Ns + "PageProps", Cell("PageWidth", width), Cell("PageHeight", height),
                            new XElement(Ns + "PageScale", new XAttribute("Unit", "IN"), pageScale.ToString("R", CultureInfo.InvariantCulture)),
                            new XElement(Ns + "DrawingScale", new XAttribute("Unit", "IN"), drawingScale.ToString("R", CultureInfo.InvariantCulture)))); break;
                    case 0x85:
                        sheet.Add(new XElement(Ns + "Line", Cell("LineWeight", chunk.Number(1)), Cell("LineColor", Color(chunk, 10)),
                            Cell("LineColorTrans", chunk.Byte(13) / 255D),
                            Cell("LinePattern", chunk.Byte(14)), Cell("BeginArrow", chunk.Byte(25)), Cell("EndArrow", chunk.Byte(26)), Cell("LineCap", chunk.Byte(27)))); break;
                    case 0x86:
                        bool indexed = Enumerable.Range(1, 4).Concat(Enumerable.Range(6, 4)).All(offset => chunk.Byte(offset) == 0);
                        sheet.Add(new XElement(Ns + "Fill",
                            Cell("FillForegnd", indexed ? chunk.Byte(0).ToString(CultureInfo.InvariantCulture) : Color(chunk, 1)),
                            Cell("FillBkgnd", indexed ? chunk.Byte(5).ToString(CultureInfo.InvariantCulture) : Color(chunk, 6)),
                            Cell("FillForegndTrans", chunk.Byte(4) / 255D), Cell("FillBkgndTrans", chunk.Byte(9) / 255D),
                            Cell("FillPattern", chunk.Byte(10)))); break;
                    case 0x87:
                        sheet.Add(new XElement(Ns + "TextBlock", Cell("LeftMargin", chunk.Number(1)), Cell("RightMargin", chunk.Number(10)),
                            Cell("TopMargin", chunk.Number(19)), Cell("BottomMargin", chunk.Number(28)), Cell("VerticalAlign", chunk.Byte(36)))); break;
                    case 0x94:
                        sheet.Add(new XElement(Ns + "Char", new XAttribute("IX", chunk.Id), Cell("Font", (ushort)(chunk.Byte(4) | chunk.Byte(5) << 8)),
                            Cell("Color", Color(chunk, 7)), Cell("Style", chunk.Byte(11) & 7), Cell("Size", chunk.Number(18)))); break;
                    case 0x6c:
                        geometry = new XElement(Ns + "Geom", new XAttribute("IX", sheet.Elements(Ns + "Geom").Count())); sheet.Add(geometry); break;
                    case 0x89:
                        geometry ??= NewGeometry(sheet);
                        geometry.Add(Cell("NoFill", chunk.Byte(0) & 1), Cell("NoLine", (chunk.Byte(0) >> 1) & 1), Cell("NoShow", (chunk.Byte(0) >> 2) & 1)); break;
                    case 0x8a: case 0x8b: case 0x8c: case 0x8f:
                        geometry ??= NewGeometry(sheet);
                        string rowName = chunk.Type switch { 0x8a => "MoveTo", 0x8b => "LineTo", 0x8c => "ArcTo", _ => "Ellipse" };
                        var row = new XElement(Ns + rowName, new XAttribute("IX", chunk.Id), Cell("X", chunk.Number(1)), Cell("Y", chunk.Number(10)));
                        if (chunk.Type is 0x8c or 0x8f) row.Add(Cell("A", chunk.Number(19)));
                        if (chunk.Type == 0x8f) row.Add(Cell("B", chunk.Number(28)), Cell("C", chunk.Number(37)), Cell("D", chunk.Number(46)));
                        geometry.Add(row); break;
                    case 0xa1:
                        string field;
                        uint name = chunk.U32(8);
                        if (chunk.Byte(7) == 232 && names.TryGetValue(name, out string? cached)) field = cached;
                        else { field = ""; Unmapped(chunk.Type); }
                        if (fields.ContainsKey(chunk.Id)) throw new InvalidDataException("Binary Visio text field identity is duplicated.");
                        fields.Add(chunk.Id, field); break;
                    case 0xe:
                        if (sheet.Element(Ns + "Text") != null) throw new InvalidDataException("Binary Visio text record is duplicated.");
                        sheet.Add(new XElement(Ns + "Text", Text(chunk.Text(8)))); break;
                    case 0x46: case 0x47: case 0x48: case 0x4a: case 0x4d: case 0x4e:
                    case 0x2d: case 0xc9: case 0x68: case 0x65: case 0x66: case 0x69: case 0x6a: case 0x2c:
                        break; // Header, identity, name and row-list records have no independent visual content.
                    default:
                        Unmapped(chunk.Type);
                        if (chunk.Type is 0x8d or 0x90 or 0xa5 or 0xa6 or 0xc1 or 0xc3) {
                            geometry ??= NewGeometry(sheet);
                            geometry.SetElementValue(Ns + "NoShow", "1");
                        }
                        break;
                }
            }
            if (sheet.Element(Ns + "Text") is XElement text && text.Value.IndexOf('\ufffc') >= 0) {
                var output = new StringBuilder();
                uint index = 0;
                foreach (char character in text.Value) {
                    if (character != '\ufffc') output.Append(character);
                    else { output.Append(fields.TryGetValue(index++, out string? cached) ? cached : ""); }
                }
                Text(output.ToString()); // Charge expanded cached fields as well as source text.
                text.Value = output.ToString();
                _findings.Add(new OfficeCompatibilityFinding("VSD_CACHED_TEXT_FIELDS", "Text",
                    "Cached string fields are substituted in text; unresolved or numeric fields are omitted and native field recalculation is not retained.",
                    OfficeCompatibilityState.Approximated, OfficeCompatibilitySeverity.Warning,
                    OfficeCompatibilityImpact.Semantic | OfficeCompatibilityImpact.Editability, representsLoss: true,
                    sourceLocation: kind + ":" + id));
            }
            return sheet;
        }

        private static XElement NewGeometry(XElement sheet) {
            var geometry = new XElement(Ns + "Geom", new XAttribute("IX", sheet.Elements(Ns + "Geom").Count()));
            sheet.Add(geometry); return geometry;
        }
        private static string Color(VisioBinaryChunks.Chunk chunk, int offset) =>
            "#" + chunk.Byte(offset).ToString("X2", CultureInfo.InvariantCulture) + chunk.Byte(offset + 1).ToString("X2", CultureInfo.InvariantCulture) + chunk.Byte(offset + 2).ToString("X2", CultureInfo.InvariantCulture);

        private XElement ReadColors(VisioBinaryContainer.Node node) {
            var result = new XElement(Ns + "Colors");
            int start = node.Shift;
            VisioBinaryData.Require(node.Data, start, 4);
            int count = node.Data[start + 2];
            VisioBinaryData.Require(node.Data, start + 4, count * 4);
            for (int index = 0; index < count; index++) {
                int offset = start + 4 + index * 4;
                string color = "#" + string.Concat(node.Data.Skip(offset).Take(3).Select(value => value.ToString("X2", CultureInfo.InvariantCulture)));
                result.Add(new XElement(Ns + "ColorEntry", new XAttribute("IX", index), new XAttribute("RGB", color)));
            }
            return result;
        }

        private XElement ReadFonts(VisioBinaryContainer.Node node) {
            var result = new XElement(Ns + "FaceNames");
            foreach (var font in node.Children.Where(child => child.Type == 0xd7)) {
                string name = ReadName(font, 4);
                result.Add(new XElement(Ns + "FaceName", Id(font.Id), new XAttribute("Name", name)));
            }
            return result;
        }

        private XElement ReadStyles(VisioBinaryContainer.Node node) {
            var chunks = Chunks(node);
            var styles = new XElement(Ns + "StyleSheets");
            for (int index = 0; index < chunks.Count; index++) {
                if (chunks[index].Type != 0x4a) continue;
                int end = index + 1;
                while (end < chunks.Count && chunks[end].Type != 0x4a) end++;
                XElement style = ReadSheet(chunks.GetRange(index, end - index), "StyleSheet", chunks[index].Id);
                ApplyStyleReferences(style, chunks[index]); styles.Add(style); index = end - 1;
            }
            return styles;
        }
    }
}
