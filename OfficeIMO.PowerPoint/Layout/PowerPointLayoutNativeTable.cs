using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;
using Ns = OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces;

namespace OfficeIMO.PowerPoint {
    /// <summary>Retains native table/cell identity while layout supplies frame geometry and continuation.</summary>
    internal sealed class PowerPointLayoutNativeTable {
        private const int D = Ns.DrawingML;
        private const int P = Ns.PresentationML;

        private readonly PowerPointLayoutNativeTable _root;
        private readonly IReadOnlyList<OfficeXmlNode> _rows;
        private readonly OfficeXmlNode _header;
        private readonly PowerPointPartRef _sourcePart;
        private bool _used;

        internal PowerPointLayoutNativeTable(PowerPointTable table, bool headers) {
            Table = table;
            _root = this;
            _sourcePart = table.SlidePartRef;
            var rows = table.TableNode.Children(D, "tr").ToList();
            _header = headers ? rows[0] : default;
            _rows = headers ? rows.Skip(1).ToArray() : rows.ToArray();
        }

        private PowerPointLayoutNativeTable(PowerPointLayoutNativeTable source, IReadOnlyList<OfficeXmlNode> rows) {
            Table = source.Table; _root = source._root; _header = source._header; _rows = rows; _sourcePart = source._sourcePart;
        }

        internal PowerPointTable Table { get; }
        internal IReadOnlyList<string> Headers => _header.IsNone
            ? Enumerable.Range(1, Table.ColumnCount).Select(i => "Column " + i).ToArray() : Values(_header);
        internal IReadOnlyList<IReadOnlyList<string>> Rows => _rows.Select(Values).ToArray();
        private IReadOnlyList<string> Values(OfficeXmlNode row) => row.Children(D, "tc").ToList().Select(c => Table.CellFor(c).Text).ToArray();

        internal PowerPointLayoutNativeTable Slice(IReadOnlyList<IReadOnlyList<string>> sourceRows, IReadOnlyList<IReadOnlyList<string>> selected) {
            var indices = selected.Select(row => Enumerable.Range(0, sourceRows.Count).Single(i => ReferenceEquals(sourceRows[i], row)));
            return new PowerPointLayoutNativeTable(this, indices.Select(i => _rows[i]).ToArray());
        }

        internal PowerPointTable Bind(PowerPointSlide slide, PowerPointTable generated) {
            var root = _root;
            var result = root._used ? generated : Table;
            OfficeXmlNode template = generated.TableNode;
            var defaults = template.Children(D, "tr").ToList();
            var rows = new List<OfficeXmlNode>();
            var sources = new List<PowerPointPartRef>();
            if (_header.Exists) {
                sources.Add(SourceOf(_header));
                rows.Add(root._used ? _header.Clone() : _header);
            }
            rows.AddRange(_rows);
            sources.AddRange(_rows.Select(SourceOf));
            for (int i = 0; i < rows.Count; i++) {
                var row = rows[i];
                slide.RehomeLayoutRelationships(row, sources[i]);
                row.SetAttribute("h", defaults[i].Attribute("h"));
                var cells = row.Children(D, "tc").ToList();
                var fallback = defaults[i].Children(D, "tc").ToList();
                for (int c = 0; c < cells.Count; c++) {
                    PowerPointLayoutNativeText.MergeMissing(cells[c].GetOrAddChild(D, "tcPr", ATableCellView.Order), fallback[c].Child(D, "tcPr"));
                    var paragraphs = cells[c].Child(D, "txBody").Children(D, "p").ToList();
                    var paragraphDefault = fallback[c].Child(D, "txBody").Child(D, "p");
                    foreach (var paragraph in paragraphs) {
                        PowerPointLayoutNativeText.MergeMissing(paragraph.GetOrAddChild(D, "pPr", ATextParagraphView.Order), paragraphDefault.Child(D, "pPr"));
                        var runDefault = paragraphDefault.Child(D, "r").Child(D, "rPr");
                        foreach (var run in paragraph.Children(D, "r"))
                            PowerPointLayoutNativeText.MergeMissing(run.GetOrPrependChild(D, "rPr"), runDefault);
                    }
                }
            }

            var properties = Table.TableNode.Child(D, "tblPr").Clone();
            var grid = template.Child(D, "tblGrid").Clone();
            var transform = generated.XmlElement.Child(P, "xfrm").Clone();
            foreach (var row in rows) if (row.Parent.Exists) row.Remove();
            OfficeXmlNode target = result.TableNode;
            target.RemoveChildren(D, "tr");
            PowerPointLayoutNativeText.Replace(target, D, "tblPr", properties, ATableView.Order);
            PowerPointLayoutNativeText.Replace(target, D, "tblGrid", grid, ATableView.Order);
            foreach (var row in rows) target.Append(row);
            PowerPointLayoutNativeText.Replace(result.XmlElement, P, "xfrm", transform, PGraphicalObjectFrameView.Order);
            if (!ReferenceEquals(result, generated)) {
                slide.RemoveShape(generated);
                slide.AdoptLayoutShape(result);
            }
            root._used = true;
            return result;
        }

        // The slide a row is on now, or the table's slide for a row that is not on one.
        private PowerPointPartRef SourceOf(OfficeXmlNode row) {
            PowerPointPartRef part = row.IsSdk ? PowerPointPartRef.FromSdk(OfficeSdkXml.PartOf(row.SdkElement!)) : default;
            return !part.IsNone && part.Kind == PowerPointPartKind.Slide ? part : _sourcePart;
        }
    }
}
